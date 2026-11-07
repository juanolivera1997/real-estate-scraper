#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Real Estate Market Scraper - Extracción de Datos Inmobiliarios y Análisis de Precio por m²
==========================================================================================

Script modular que:
  1. Extrae anuncios inmobiliarios (venta y alquiler) de un portal web con paginación.
  2. Limpia y normaliza los datos con pandas (precios, monedas, superficies, estado, antigüedad).
  3. Calcula métricas clave: Precio por m², precio normalizado en USD y USD/m².
  4. Elimina duplicados (mismo enlace o mismo título + ubicación).
  5. Genera un reporte corporativo en Excel (.xlsx) con OpenPyXL.

Incluye un servidor DEMO local (ThreadingHTTPServer) con 5 páginas de catálogo simulado
con datos "semi-sucios", para que el proyecto sea 100 % reproducible sin depender de
sitios externos.

Uso rápido:
    python scraper_inmobiliario.py --demo
    python scraper_inmobiliario.py --url "https://portal.com/listado" --selectors selectores.json

Autor: Juan Pablo (Portafolio Python · Web Scraping · Data Engineering)
Licencia: MIT
"""

from __future__ import annotations

import argparse
import json
import logging
import random
import re
import sys
import threading
import time
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Optional
from urllib.parse import parse_qs, urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ════════════════════════════════════════════════════════════════════════════════
# 1. CONFIGURACIÓN
# ════════════════════════════════════════════════════════════════════════════════

LOG = logging.getLogger("real_estate_scraper")

DEFAULT_OUTPUT = "reporte_mercado_inmobiliario.xlsx"
NA_VALUE = "N/D"
CURRENT_YEAR = datetime.now().year

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/130.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_6) AppleWebKit/605.1.15 (KHTML, like Gecko) "
    "Version/18.0 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64; rv:131.0) Gecko/20100101 Firefox/131.0",
]

# Selectores CSS por defecto (coinciden con el servidor DEMO).
# Para un portal real, se pueden sobrescribir con --selectors archivo.json
DEFAULT_SELECTORS: dict[str, str] = {
    "card": "article.listing-card",
    "title": "h2.listing-title",
    "link": "h2.listing-title a",
    "operation": "span.badge-operation",
    "property_type": "span.property-type",
    "location": "p.listing-location",
    "price": "div.listing-price",
    "area": "li.feature-area",
    "bedrooms": "li.feature-bedrooms",
    "bathrooms": "li.feature-bathrooms",
    "condition": "li.feature-condition",
    "age": "li.feature-age",
    "next_page": "a[rel='next']",
}


@dataclass
class ScraperConfig:
    """Parámetros de ejecución del scraper."""

    start_url: str
    max_pages: int = 5
    delay_min: float = 1.0
    delay_max: float = 2.5
    timeout: int = 15
    fx_rate: float = 40.0  # Moneda local por 1 USD (configurable con --fx-rate)
    output: str = DEFAULT_OUTPUT
    selectors: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_SELECTORS))
    bypass_proxy: bool = False  # True en modo DEMO (localhost)


@dataclass
class RunStats:
    """Métricas de la ejecución para el resumen final."""

    pages_scraped: int = 0
    raw_listings: int = 0
    discarded_invalid: int = 0
    duplicates: int = 0
    final_listings: int = 0
    errors: int = 0


# ════════════════════════════════════════════════════════════════════════════════
# 2. SERVIDOR DEMO (catálogo inmobiliario simulado con datos semi-sucios)
# ════════════════════════════════════════════════════════════════════════════════

class DemoCatalog:
    """Genera de forma determinística 5 páginas de anuncios con datos 'sucios'."""

    PAGES = 5
    PER_PAGE = 12

    LOCATIONS = [
        "Pocitos, Montevideo", "Punta Carretas, Montevideo", "Carrasco, Montevideo",
        "Cordón, Montevideo", "Centro, Montevideo", "Malvín, Montevideo",
        "Buceo, Montevideo", "Parque Rodó, Montevideo", "Prado, Montevideo",
        "Península, Punta del Este", "Ciudad de la Costa, Canelones",
    ]
    TYPE_LABELS = {
        "Apartamento": ["Apartamento", "Apto.", " APARTAMENTO ", "Apto"],
        "Casa": ["Casa", " casa", "CASA "],
        "Terreno": ["Terreno", "Solar / Terreno", " terreno"],
        "Oficina": ["Oficina", "Local / Oficina", "OFICINA"],
    }
    CONDITIONS = ["A estrenar", "Excelente estado", " Buen estado ", "Muy buen estado",
                  "A reciclar", "En construcción", "buen estado"]

    def __init__(self, seed: int = 42) -> None:
        self.rng = random.Random(seed)
        self.listings = self._build()

    # --- Formateadores "sucios" -------------------------------------------------
    def _dirty_price(self, amount: int, currency: str, operation: str) -> str:
        if self.rng.random() < 0.04:
            return self.rng.choice(["Consultar", "Precio a convenir", " -- "])
        if currency == "USD":
            fmt = self.rng.choice([
                "U$S {:,.0f}", "USD {:,.0f}", "US$ {:,.0f} ", "  u$s{:,.0f}", "USD {:,.0f}.-",
            ])
            txt = fmt.format(amount)
            if self.rng.random() < 0.6:  # Separador de miles latino (punto)
                txt = txt.replace(",", ".")
        else:
            fmt = self.rng.choice(["$U {:,.0f}", "$ {:,.0f}", "UYU {:,.0f}", "$ {:,.0f} /mes"])
            txt = fmt.format(amount).replace(",", ".")
        if operation == "Alquiler" and "/mes" not in txt and self.rng.random() < 0.3:
            txt += " por mes"
        return txt

    def _dirty_area(self, area: float) -> str:
        fmt = self.rng.choice(["{} m²", "{}m2", " Sup. total: {} mts² ", "{} m2 totales", "{} M2"])
        if area >= 1000:
            val = f"{area:,.0f}".replace(",", ".")
        elif self.rng.random() < 0.25:
            val = f"{area + 0.5:.1f}".replace(".", ",")
        else:
            val = f"{area:.0f}"
        return fmt.format(val)

    def _dirty_rooms(self, n: Optional[int], kind: str) -> Optional[str]:
        if n is None:
            return None
        if kind == "bed":
            if n == 0:
                return "Monoambiente"
            return self.rng.choice([f"{n} dorm.", f"{n} dormitorios", f" {n} Dormitorio(s)"])
        return self.rng.choice([f"{n} baño" if n == 1 else f"{n} baños", f"{n} bñ", f" {n} Baños "])

    def _dirty_age(self, year: Optional[int], condition: str) -> Optional[str]:
        if year is None:
            return None
        if "estrenar" in condition.lower():
            return "A estrenar"
        if "construcción" in condition.lower():
            return self.rng.choice(["En obra", f"Entrega {CURRENT_YEAR + 1}"])
        return self.rng.choice([
            f"Año {year}", f"Construido en {year}", f"Antigüedad: {CURRENT_YEAR - year} años",
            f" {CURRENT_YEAR - year} años ",
        ])

    # --- Construcción del catálogo ------------------------------------------------
    def _build(self) -> list[dict[str, Any]]:
        listings: list[dict[str, Any]] = []
        total = self.PAGES * self.PER_PAGE
        for i in range(1, total + 1):
            ptype = self.rng.choices(
                ["Apartamento", "Casa", "Terreno", "Oficina"], weights=[50, 25, 10, 15]
            )[0]
            operation = "Venta" if (ptype == "Terreno" or self.rng.random() < 0.6) else "Alquiler"
            location = self.rng.choice(self.LOCATIONS)
            premium = 1.6 if any(z in location for z in ("Carrasco", "Punta", "Península")) else 1.0

            area = {
                "Apartamento": self.rng.randint(35, 160),
                "Casa": self.rng.randint(90, 380),
                "Terreno": self.rng.randint(300, 2500),
                "Oficina": self.rng.randint(25, 220),
            }[ptype]

            if operation == "Venta":
                usd_m2 = {"Apartamento": 2900, "Casa": 2100, "Terreno": 230, "Oficina": 2400}[ptype]
                amount = round(area * usd_m2 * premium * self.rng.uniform(0.75, 1.3), -3)
                currency = "USD"
            else:
                uyu_m2 = {"Apartamento": 420, "Casa": 330, "Oficina": 460}[ptype]
                amount = round(area * uyu_m2 * premium * self.rng.uniform(0.8, 1.25), -2)
                currency = "USD" if self.rng.random() < 0.15 else "UYU"
                if currency == "USD":
                    amount = round(amount / 40, -1)

            beds = baths = year = None
            condition = None
            if ptype in ("Apartamento", "Casa"):
                beds = max(0, min(5, area // 40 - (0 if ptype == "Apartamento" else 1)))
                baths = max(1, min(4, (beds or 1) // 2 + self.rng.randint(0, 1)))
            if ptype == "Oficina":
                baths = self.rng.randint(1, 2)
            if ptype != "Terreno":
                condition = self.rng.choice(self.CONDITIONS)
                cond_key = condition.lower()
                if "construcción" in cond_key or "estrenar" in cond_key:
                    year = CURRENT_YEAR
                elif "reciclar" in cond_key:
                    year = self.rng.randint(1940, 1990)
                else:
                    year = self.rng.randint(1965, CURRENT_YEAR - 1)

            feminine = ptype in ("Casa", "Oficina")
            adjective = self.rng.choice(
                ["Luminosa", "Amplia", "Moderna", "Excelente", "Hermosa", "Céntrica"] if feminine
                else ["Luminoso", "Amplio", "Moderno", "Excelente", "Hermoso", "Céntrico"])
            if self.rng.random() < 0.15:
                adjective = "Oportunidad:"
            title = f"{adjective} {ptype.lower()} en {location.split(',')[0]}"
            if beds and ptype in ("Apartamento", "Casa"):
                title += f" - {beds} dormitorio{'s' if beds > 1 else ''}"
            elif beds == 0:
                title = f"Monoambiente en {location.split(',')[0]}"

            listings.append({
                "id": 1000 + i,
                "title": ("  " + title + " ") if self.rng.random() < 0.2 else title,
                "operation": self.rng.choice(
                    [operation, operation.upper(), f"En {operation}", f" {operation.lower()} "]),
                "ptype": self.rng.choice(self.TYPE_LABELS[ptype]),
                "location": location if self.rng.random() > 0.15 else f"  {location.upper()} ",
                "price": self._dirty_price(int(amount), currency, operation),
                "area": self._dirty_area(area) if self.rng.random() > 0.05 else None,
                "beds": self._dirty_rooms(beds, "bed"),
                "baths": self._dirty_rooms(baths, "bath"),
                "condition": condition,
                "age": self._dirty_age(year, condition or ""),
                "href": f"/propiedad/{1000 + i}",
            })

        # Duplicados intencionales: mismo enlace repetido en otra página...
        for src, dst in [(3, 20), (15, 31), (27, 44), (40, 55)]:
            listings[dst] = dict(listings[src])
        # ...y misma publicación con distinto enlace (título + ubicación coinciden)
        for src, dst in [(8, 36), (22, 58)]:
            dup = dict(listings[src])
            dup["href"] = f"/propiedad/{dup['id']}-republicado"
            dup["title"] = dup["title"].upper()
            listings[dst] = dup
        # Tarjetas inválidas (banners publicitarios sin título ni enlace)
        for idx in (11, 47):
            listings[idx] = {"invalid": True}
        return listings

    def page(self, n: int) -> list[dict[str, Any]]:
        start = (n - 1) * self.PER_PAGE
        return self.listings[start:start + self.PER_PAGE]


def _render_card(item: dict[str, Any]) -> str:
    if item.get("invalid"):
        return ('<article class="listing-card sponsored"><div class="listing-price">'
                '¡Tasá tu propiedad gratis!</div></article>')

    def li(cls: str, value: Optional[str]) -> str:
        return f'<li class="{cls}">{value}</li>' if value else ""

    return f"""
    <article class="listing-card" data-id="{item['id']}">
      <span class="badge-operation">{item['operation']}</span>
      <span class="property-type">{item['ptype']}</span>
      <h2 class="listing-title"><a href="{item['href']}">{item['title']}</a></h2>
      <p class="listing-location">{item['location']}</p>
      <div class="listing-price">{item['price']}</div>
      <ul class="listing-features">
        {li('feature-area', item['area'])}{li('feature-bedrooms', item['beds'])}
        {li('feature-bathrooms', item['baths'])}{li('feature-condition', item['condition'])}
        {li('feature-age', item['age'])}
      </ul>
    </article>"""


def _render_page(catalog: DemoCatalog, n: int) -> str:
    cards = "\n".join(_render_card(it) for it in catalog.page(n))
    prev_link = f'<a rel="prev" href="/listado?page={n - 1}">« Anterior</a>' if n > 1 else ""
    next_link = (f'<a rel="next" href="/listado?page={n + 1}">Siguiente »</a>'
                 if n < catalog.PAGES else "")
    return f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8"><title>DemoPropiedades · Página {n}</title></head>
<body>
  <header><h1>DemoPropiedades - Catálogo de inmuebles</h1></header>
  <main class="results">{cards}</main>
  <nav class="pagination">{prev_link} <span>Página {n} de {catalog.PAGES}</span> {next_link}</nav>
</body></html>"""


def make_demo_handler(catalog: DemoCatalog) -> type[BaseHTTPRequestHandler]:
    class DemoHandler(BaseHTTPRequestHandler):
        def do_GET(self) -> None:  # noqa: N802 (nombre impuesto por http.server)
            parsed = urlparse(self.path)
            if parsed.path in ("/", "/listado"):
                try:
                    n = int(parse_qs(parsed.query).get("page", ["1"])[0])
                except ValueError:
                    n = 1
                if 1 <= n <= catalog.PAGES:
                    return self._send(200, _render_page(catalog, n))
            elif parsed.path.startswith("/propiedad/"):
                pid = parsed.path.rsplit("/", 1)[-1]
                return self._send(200, f"<html><body><h1>Ficha de la propiedad {pid}</h1>"
                                       f"<p>Página de detalle (DEMO).</p></body></html>")
            self._send(404, "<h1>404 - No encontrado</h1>")

        def _send(self, code: int, body: str) -> None:
            data = body.encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def log_message(self, fmt: str, *args: Any) -> None:  # Silenciar logs del servidor
            LOG.debug("[DEMO-SERVER] " + fmt, *args)

    return DemoHandler


class DemoServer:
    """Context manager que levanta el servidor DEMO en un hilo en segundo plano."""

    def __init__(self, host: str = "127.0.0.1", port: int = 0) -> None:
        self.catalog = DemoCatalog()
        self.httpd = ThreadingHTTPServer((host, port), make_demo_handler(self.catalog))
        self.thread = threading.Thread(target=self.httpd.serve_forever, daemon=True)

    @property
    def base_url(self) -> str:
        host, port = self.httpd.server_address[:2]
        return f"http://{host}:{port}/listado?page=1"

    def __enter__(self) -> "DemoServer":
        self.thread.start()
        LOG.info("Servidor DEMO activo en %s", self.base_url)
        return self

    def __exit__(self, *exc: Any) -> None:
        self.httpd.shutdown()
        self.httpd.server_close()
        LOG.info("Servidor DEMO detenido.")


# ════════════════════════════════════════════════════════════════════════════════
# 3. CAPA DE EXTRACCIÓN (HTTP + parsing HTML)
# ════════════════════════════════════════════════════════════════════════════════

class RealEstateScraper:
    """Descarga páginas de listado, sigue la paginación y extrae cada anuncio."""

    def __init__(self, config: ScraperConfig, stats: RunStats) -> None:
        self.cfg = config
        self.stats = stats
        self.session = self._build_session()

    def _build_session(self) -> requests.Session:
        session = requests.Session()
        session.headers.update({
            "User-Agent": random.choice(USER_AGENTS),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "es-UY,es;q=0.9,en;q=0.8",
            "Accept-Encoding": "gzip, deflate",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        })
        retry = Retry(total=3, backoff_factor=1.5, status_forcelist=(429, 500, 502, 503, 504),
                      allowed_methods=("GET",), respect_retry_after_header=True)
        adapter = HTTPAdapter(max_retries=retry)
        session.mount("http://", adapter)
        session.mount("https://", adapter)
        if self.cfg.bypass_proxy:
            session.trust_env = False  # localhost: ignorar proxies del sistema
        return session

    def _polite_pause(self) -> None:
        """Pausa aleatoria (jitter) entre solicitudes para no saturar el servidor."""
        delay = random.uniform(self.cfg.delay_min, self.cfg.delay_max)
        LOG.debug("Pausa de cortesía: %.2fs", delay)
        time.sleep(delay)

    def fetch(self, url: str) -> Optional[BeautifulSoup]:
        try:
            resp = self.session.get(url, timeout=self.cfg.timeout)
            resp.raise_for_status()
            resp.encoding = resp.encoding or resp.apparent_encoding
            return BeautifulSoup(resp.text, "html.parser")
        except requests.RequestException as exc:
            self.stats.errors += 1
            LOG.error("Error al descargar %s → %s", url, exc)
            return None

    # --- Extracción segura por campo -------------------------------------------
    @staticmethod
    def _text(node: Tag, selector: str) -> Optional[str]:
        try:
            el = node.select_one(selector) if selector else None
            if el is None:
                return None
            txt = el.get_text(" ", strip=True)
            return txt or None
        except Exception as exc:  # selector inválido u otro problema puntual
            LOG.debug("Campo no disponible (%s): %s", selector, exc)
            return None

    def _link(self, node: Tag, page_url: str) -> Optional[str]:
        try:
            el = node.select_one(self.cfg.selectors["link"])
            href = el.get("href") if el else None
            return urljoin(page_url, href.strip()) if href else None
        except Exception as exc:
            LOG.debug("Enlace no disponible: %s", exc)
            return None

    def parse_listing(self, card: Tag, page_url: str) -> dict[str, Optional[str]]:
        s = self.cfg.selectors
        return {
            "titulo": self._text(card, s["title"]),
            "operacion": self._text(card, s["operation"]),
            "tipo": self._text(card, s["property_type"]),
            "ubicacion": self._text(card, s["location"]),
            "precio_raw": self._text(card, s["price"]),
            "superficie_raw": self._text(card, s["area"]),
            "dormitorios_raw": self._text(card, s["bedrooms"]),
            "banos_raw": self._text(card, s["bathrooms"]),
            "estado_raw": self._text(card, s["condition"]),
            "antiguedad_raw": self._text(card, s["age"]),
            "enlace": self._link(card, page_url),
        }

    def _next_page_url(self, soup: BeautifulSoup, current: str) -> Optional[str]:
        try:
            el = soup.select_one(self.cfg.selectors["next_page"])
            if el and el.get("href"):
                return urljoin(current, el["href"])
        except Exception as exc:
            LOG.debug("No se pudo resolver la página siguiente: %s", exc)
        return None

    def run(self) -> list[dict[str, Optional[str]]]:
        results: list[dict[str, Optional[str]]] = []
        url: Optional[str] = self.cfg.start_url
        visited: set[str] = set()

        while url and url not in visited and self.stats.pages_scraped < self.cfg.max_pages:
            if visited:
                self._polite_pause()
            visited.add(url)
            LOG.info("Página %d → %s", self.stats.pages_scraped + 1, url)
            soup = self.fetch(url)
            if soup is None:
                break
            self.stats.pages_scraped += 1

            cards = soup.select(self.cfg.selectors["card"])
            if not cards:
                LOG.warning("No se encontraron tarjetas con el selector '%s'.",
                            self.cfg.selectors["card"])
            for card in cards:
                try:
                    item = self.parse_listing(card, url)
                    item["pagina"] = str(self.stats.pages_scraped)
                    results.append(item)
                except Exception as exc:  # Nunca detener el proceso por un anuncio
                    self.stats.errors += 1
                    LOG.warning("Anuncio omitido por error inesperado: %s", exc)
            LOG.info("   %d anuncios encontrados en la página.", len(cards))
            url = self._next_page_url(soup, url)

        self.stats.raw_listings = len(results)
        return results


# ════════════════════════════════════════════════════════════════════════════════
# 4. LIMPIEZA Y NORMALIZACIÓN (funciones puras, fáciles de testear)
# ════════════════════════════════════════════════════════════════════════════════

def _strip_accents(text: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))


def normalize_key(text: Optional[str]) -> str:
    """Normaliza texto para comparar duplicados (sin tildes, minúsculas, espacios simples)."""
    if not isinstance(text, str) or not text:
        return ""
    return re.sub(r"\s+", " ", _strip_accents(str(text)).lower()).strip()


def clean_text(text: Optional[str]) -> Optional[str]:
    if not isinstance(text, str):
        return None
    cleaned = re.sub(r"\s+", " ", str(text)).strip()
    return cleaned or None


def parse_number(raw: Optional[str]) -> Optional[float]:
    """
    Convierte cadenas numéricas con formatos mixtos a float.
      '185.000' → 185000 · '1,250,000' → 1250000 · '120,5' → 120.5
      '1.234,56' → 1234.56 · '1,234.56' → 1234.56
    """
    if not isinstance(raw, str) or not raw:
        return None
    match = re.search(r"\d[\d.,]*", str(raw))
    if not match:
        return None
    num = match.group(0).rstrip(".,")
    has_dot, has_comma = "." in num, "," in num

    if has_dot and has_comma:  # El último separador es el decimal
        if num.rfind(",") > num.rfind("."):
            num = num.replace(".", "").replace(",", ".")
        else:
            num = num.replace(",", "")
    elif has_dot or has_comma:
        sep = "." if has_dot else ","
        parts = num.split(sep)
        # Separador de miles si se repite o si le siguen exactamente 3 dígitos
        if len(parts) > 2 or len(parts[-1]) == 3:
            num = num.replace(sep, "")
        else:
            num = num.replace(sep, ".")
    try:
        return float(num)
    except ValueError:
        return None


CURRENCY_PATTERNS = [
    (re.compile(r"(u\$s|us\$|usd|u\$d|dolares|dólares)", re.I), "USD"),
    (re.compile(r"(\$u|uyu|pesos|\$)", re.I), "UYU"),
    (re.compile(r"(€|eur)", re.I), "EUR"),
]


def parse_price(raw: Optional[str]) -> tuple[Optional[float], Optional[str]]:
    """Devuelve (monto, moneda). 'Consultar' u otros textos → (None, None)."""
    if not isinstance(raw, str) or not raw:
        return None, None
    amount = parse_number(raw)
    if amount is None or amount <= 0:
        return None, None
    currency = next((code for rx, code in CURRENCY_PATTERNS if rx.search(raw)), None)
    return amount, currency


def parse_int(raw: Optional[str]) -> Optional[int]:
    if not isinstance(raw, str) or not raw:
        return None
    if "monoambiente" in raw.lower():
        return 0
    m = re.search(r"\d+", raw)
    return int(m.group(0)) if m else None


OPERATION_MAP = {"venta": "Venta", "vende": "Venta", "alquiler": "Alquiler",
                 "alquila": "Alquiler", "arriendo": "Alquiler", "renta": "Alquiler"}
TYPE_MAP = [("apart", "Apartamento"), ("apto", "Apartamento"), ("depto", "Apartamento"),
            ("departamento", "Apartamento"), ("casa", "Casa"), ("terreno", "Terreno"),
            ("solar", "Terreno"), ("lote", "Terreno"), ("oficina", "Oficina"),
            ("local", "Oficina")]
CONDITION_MAP = [("estrenar", "A estrenar"), ("construccion", "En construcción"),
                 ("reciclar", "A reciclar"), ("excelente", "Excelente"),
                 ("muy buen", "Muy bueno"), ("buen", "Bueno")]


def normalize_operation(raw: Optional[str], price_raw: Optional[str] = None) -> Optional[str]:
    key = normalize_key(raw)
    for token, label in OPERATION_MAP.items():
        if token in key:
            return label
    # Inferencia: los alquileres suelen expresarse "por mes"
    if price_raw and re.search(r"/\s*mes|por mes|mensual", price_raw, re.I):
        return "Alquiler"
    return None


def normalize_type(raw: Optional[str], title: Optional[str] = None) -> Optional[str]:
    for source in (raw, title):
        key = normalize_key(source)
        for token, label in TYPE_MAP:
            if token in key:
                return label
    return None


def normalize_condition(raw: Optional[str]) -> Optional[str]:
    key = normalize_key(raw)
    for token, label in CONDITION_MAP:
        if token in key:
            return label
    return clean_text(raw).title() if raw else None


def parse_age(raw: Optional[str], condition: Optional[str] = None) -> Optional[int]:
    """Antigüedad en años a partir de 'Año 1998', 'Antigüedad: 15 años', 'A estrenar'…"""
    key = normalize_key(raw)
    if (any(t in key for t in ("estrenar", "en obra", "entrega", "construccion"))
            or (condition in ("A estrenar", "En construcción") and not key)):
        return 0
    if not key:
        return None
    year = re.search(r"\b(18|19|20)\d{2}\b", key)
    if year:
        return max(0, CURRENT_YEAR - int(year.group(0)))
    years = re.search(r"(\d{1,3})\s*anos?", key)
    return int(years.group(1)) if years else None


SMALL_WORDS = {"de", "del", "la", "las", "el", "los", "y", "en"}


def smart_title(text: str) -> str:
    """Capitaliza solo textos en MAYÚSCULAS/minúsculas, respetando conectores ('de', 'la')."""
    text = text.strip()
    if not (text.isupper() or text.islower()):
        return text
    words = text.lower().split()
    return " ".join(w if (i and w in SMALL_WORDS) else w.capitalize() for i, w in enumerate(words))


def split_location(raw: Optional[str]) -> tuple[Optional[str], Optional[str]]:
    text = clean_text(raw)
    if not text:
        return None, None
    parts = [smart_title(p) for p in text.split(",") if p.strip()]
    if len(parts) == 1:
        return parts[0], None
    return parts[0], parts[-1]


# ════════════════════════════════════════════════════════════════════════════════
# 5. INGENIERÍA DE DATOS CON PANDAS
# ════════════════════════════════════════════════════════════════════════════════

REPORT_COLUMNS = [
    "Título", "Operación", "Tipo de Propiedad", "Barrio", "Ciudad", "Moneda", "Precio",
    "Superficie (m²)", "Precio por m²", "Precio USD", "USD por m²", "Dormitorios", "Baños",
    "Estado", "Antigüedad (años)", "Enlace", "Página", "Fecha de Extracción",
]
NUMERIC_COLUMNS = ["Precio", "Superficie (m²)", "Precio por m²", "Precio USD", "USD por m²",
                   "Dormitorios", "Baños", "Antigüedad (años)"]


def build_dataframe(records: list[dict[str, Optional[str]]], cfg: ScraperConfig,
                    stats: RunStats) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Devuelve (df_reporte, df_numerico). df_numerico conserva NaN para estadísticas."""
    df = pd.DataFrame(records)
    if df.empty:
        return pd.DataFrame(columns=REPORT_COLUMNS), pd.DataFrame(columns=REPORT_COLUMNS)
    # Columnas crudas como objeto Python con None (evita NaN float en parsers de texto)
    df = df.astype(object).where(df.notna(), None)

    # 5.1 Descartar registros inutilizables (sin título y sin enlace, p. ej. banners)
    df["titulo"] = df["titulo"].map(clean_text)
    valid = df["titulo"].notna() | df["enlace"].notna()
    stats.discarded_invalid = int((~valid).sum())
    df = df[valid].copy()

    # 5.2 Limpieza y normalización de campos
    prices = df["precio_raw"].map(parse_price)
    df["Precio"] = prices.map(lambda t: t[0])
    df["Moneda"] = prices.map(lambda t: t[1])
    df["Superficie (m²)"] = df["superficie_raw"].map(parse_number)
    df["Operación"] = [normalize_operation(o, p) for o, p in zip(df["operacion"], df["precio_raw"])]
    df["Tipo de Propiedad"] = [normalize_type(t, ti) for t, ti in zip(df["tipo"], df["titulo"])]
    loc = df["ubicacion"].map(split_location)
    df["Barrio"] = loc.map(lambda t: t[0])
    df["Ciudad"] = loc.map(lambda t: t[1])
    df["Dormitorios"] = df["dormitorios_raw"].map(parse_int)
    df["Baños"] = df["banos_raw"].map(parse_int)
    df["Estado"] = df["estado_raw"].map(normalize_condition)
    df["Antigüedad (años)"] = [parse_age(a, c) for a, c in zip(df["antiguedad_raw"], df["Estado"])]
    df["Título"] = df["titulo"]
    df["Enlace"] = df["enlace"]
    df["Página"] = pd.to_numeric(df["pagina"], errors="coerce")
    df["Fecha de Extracción"] = datetime.now().strftime("%Y-%m-%d %H:%M")

    # 5.3 Eliminación de duplicados: (a) mismo enlace, (b) mismo título + ubicación
    before = len(df)
    df["_k_link"] = df["Enlace"].fillna("").str.lower().str.rstrip("/")
    df["_k_tl"] = df["Título"].map(normalize_key) + "|" + df["ubicacion"].map(normalize_key)
    link_dup = df["_k_link"].ne("") & df.duplicated("_k_link", keep="first")
    df = df[~link_dup]
    df = df[~df.duplicated("_k_tl", keep="first")]
    stats.duplicates = before - len(df)

    # 5.4 Métricas clave
    area = df["Superficie (m²)"].where(df["Superficie (m²)"] > 0)
    df["Precio por m²"] = (df["Precio"] / area).round(2)
    fx = {"USD": 1.0, "UYU": 1 / cfg.fx_rate}
    df["Precio USD"] = (df["Precio"] * df["Moneda"].map(fx)).round(2)
    df["USD por m²"] = (df["Precio USD"] / area).round(2)

    df_num = df[REPORT_COLUMNS].sort_values(
        ["Operación", "Tipo de Propiedad", "USD por m²"], na_position="last").reset_index(drop=True)
    stats.final_listings = len(df_num)

    # 5.5 Sustitución de nulos por "N/D" (en una copia de presentación)
    df_report = df_num.astype(object).where(df_num.notna(), NA_VALUE)
    return df_report, df_num


def build_summary(df_num: pd.DataFrame) -> pd.DataFrame:
    """Resumen por segmento de mercado (operación · tipo · moneda)."""
    valid = df_num.dropna(subset=["Precio por m²"])
    if valid.empty:
        return pd.DataFrame()
    g = valid.groupby(["Operación", "Tipo de Propiedad", "Moneda"], dropna=False)
    summary = g.agg(**{
        "Anuncios": ("Precio", "size"),
        "Precio Promedio": ("Precio", "mean"),
        "Superficie Promedio (m²)": ("Superficie (m²)", "mean"),
        "Precio por m² Promedio": ("Precio por m²", "mean"),
        "Precio por m² Mediana": ("Precio por m²", "median"),
        "Precio por m² Mínimo": ("Precio por m²", "min"),
        "Precio por m² Máximo": ("Precio por m²", "max"),
        "USD por m² Promedio": ("USD por m²", "mean"),
    }).round(2).reset_index()
    return summary.astype(object).where(summary.notna(), NA_VALUE)


# ════════════════════════════════════════════════════════════════════════════════
# 6. REPORTE EXCEL CORPORATIVO (OpenPyXL)
# ════════════════════════════════════════════════════════════════════════════════

HEADER_COLOR = "1E4620"   # Verde oscuro corporativo (alternativa: "1B365D" Azul Petróleo)
FONT_NAME = "Aptos"       # Alternativa: "Segoe UI"
FMT_THOUSANDS = "#,##0"
FMT_THOUSANDS_DEC = "#,##0.00"

COLUMN_FORMATS = {
    "Precio": FMT_THOUSANDS, "Precio USD": FMT_THOUSANDS, "Precio por m²": FMT_THOUSANDS_DEC,
    "USD por m²": FMT_THOUSANDS_DEC, "Superficie (m²)": "#,##0.0",
    "Precio Promedio": FMT_THOUSANDS, "Superficie Promedio (m²)": "#,##0.0",
    "Precio por m² Promedio": FMT_THOUSANDS_DEC, "Precio por m² Mediana": FMT_THOUSANDS_DEC,
    "Precio por m² Mínimo": FMT_THOUSANDS_DEC, "Precio por m² Máximo": FMT_THOUSANDS_DEC,
    "USD por m² Promedio": FMT_THOUSANDS_DEC,
}


def _style_sheet(ws, header_color: str = HEADER_COLOR) -> None:
    header_fill = PatternFill("solid", start_color=header_color, end_color=header_color)
    header_font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    body_font = Font(name=FONT_NAME, size=10)
    link_font = Font(name=FONT_NAME, size=10, color="0563C1", underline="single")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    zebra = PatternFill("solid", start_color="F2F7F2", end_color="F2F7F2")

    headers = [c.value for c in ws[1]]
    for cell in ws[1]:
        cell.fill, cell.font, cell.border = header_fill, header_font, border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for col_idx, cell in enumerate(row):
            name = headers[col_idx]
            cell.font, cell.border = body_font, border
            if row_idx % 2 == 0:
                cell.fill = zebra
            if name in COLUMN_FORMATS and isinstance(cell.value, (int, float)):
                cell.number_format = COLUMN_FORMATS[name]
                cell.alignment = Alignment(horizontal="right")
            elif cell.value == NA_VALUE:
                cell.alignment = Alignment(horizontal="center")
                cell.font = Font(name=FONT_NAME, size=10, color="808080", italic=True)
            if name == "Enlace" and isinstance(cell.value, str) and cell.value.startswith("http"):
                cell.hyperlink = cell.value
                cell.font = link_font

    # Auto-ajuste de ancho de columnas
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        longest = 0
        for cell in col_cells:
            value = cell.value
            if isinstance(value, float):
                text = f"{value:,.2f}"
            else:
                text = str(value) if value is not None else ""
            longest = max(longest, len(text))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(longest + 3, 10), 60)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.zoomScale = 100
    # Impresión: horizontal, ajustado al ancho de página
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.print_title_rows = "1:1"


def export_excel(df_report: pd.DataFrame, df_summary: pd.DataFrame, output: str) -> Path:
    path = Path(output).resolve()
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df_report.to_excel(writer, sheet_name="Propiedades", index=False)
        if not df_summary.empty:
            df_summary.to_excel(writer, sheet_name="Resumen por Segmento", index=False)

    wb = load_workbook(path)
    for ws in wb.worksheets:
        _style_sheet(ws)
    wb.properties.title = "Reporte de Mercado Inmobiliario"
    wb.properties.creator = "Real Estate Market Scraper"
    wb.save(path)
    return path


# ════════════════════════════════════════════════════════════════════════════════
# 7. RESUMEN EN TERMINAL
# ════════════════════════════════════════════════════════════════════════════════

def print_summary(stats: RunStats, df_num: pd.DataFrame, path: Optional[Path],
                  cfg: ScraperConfig, elapsed: float) -> None:
    line = "═" * 66
    print(f"\n{line}\n  📊  RESUMEN DE EJECUCIÓN · REAL ESTATE MARKET SCRAPER\n{line}")
    print(f"  Páginas procesadas ............... {stats.pages_scraped}")
    print(f"  Anuncios extraídos (brutos) ...... {stats.raw_listings}")
    print(f"  Descartados (inválidos) .......... {stats.discarded_invalid}")
    print(f"  Duplicados eliminados ............ {stats.duplicates}")
    print(f"  Total consolidado ................ {stats.final_listings}")
    print(f"  Errores de red/parsing ........... {stats.errors}")

    if not df_num.empty:
        no_price = int(df_num["Precio"].isna().sum())
        print(f"  Sin precio publicado ............. {no_price}")
        # El promedio se calcula por operación: mezclar venta y alquiler distorsiona la métrica
        for op, label in (("Venta", "Venta"), ("Alquiler", "Alquiler/mes")):
            serie = df_num.loc[df_num["Operación"] == op, "USD por m²"].dropna()
            value = f"US$ {serie.mean():,.2f}" if not serie.empty else NA_VALUE
            text = f"Precio prom. m² · {label} "
            print(f"  {text:.<34} {value:<14} ({len(serie)} anuncios)")
        print(f"  (Normalizado a USD con TC {cfg.fx_rate:g} moneda local/USD)")
        print("\n  Precio promedio por m² por segmento (moneda original):")
        seg = (df_num.dropna(subset=["Precio por m²"])
               .groupby(["Operación", "Moneda"])["Precio por m²"].agg(["mean", "count"]))
        for (op, cur), row in seg.iterrows():
            n = int(row["count"])
            print(f"    • {op:<9} {cur:<4} {row['mean']:>12,.2f} /m²   ({n} anuncio{'s' if n != 1 else ''})")
    print(f"\n  Tiempo total ..................... {elapsed:.1f} s")
    if path:
        print(f"  Reporte Excel .................... {path}")
    print(line)


# ════════════════════════════════════════════════════════════════════════════════
# 8. CLI / ORQUESTACIÓN
# ════════════════════════════════════════════════════════════════════════════════

def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Real Estate Market Scraper · Extracción y análisis de precio por m².",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    src = p.add_mutually_exclusive_group(required=True)
    src.add_argument("--demo", action="store_true", help="Usar el servidor DEMO local integrado.")
    src.add_argument("--url", help="URL de la primera página de listado de un portal real.")
    p.add_argument("--selectors", help="Archivo JSON con selectores CSS personalizados.")
    p.add_argument("--max-pages", type=int, default=5, help="Máximo de páginas a recorrer.")
    p.add_argument("--delay-min", type=float, default=None, help="Pausa mínima entre requests (s).")
    p.add_argument("--delay-max", type=float, default=None, help="Pausa máxima entre requests (s).")
    p.add_argument("--fx-rate", type=float, default=40.0,
                   help="Tipo de cambio: moneda local por 1 USD (para normalizar USD/m²).")
    p.add_argument("--output", default=DEFAULT_OUTPUT, help="Ruta del archivo Excel de salida.")
    p.add_argument("--port", type=int, default=0, help="Puerto del servidor DEMO (0 = automático).")
    p.add_argument("--verbose", "-v", action="store_true", help="Logs detallados (DEBUG).")
    return p.parse_args(argv)


def load_selectors(path: Optional[str]) -> dict[str, str]:
    selectors = dict(DEFAULT_SELECTORS)
    if path:
        with open(path, encoding="utf-8") as fh:
            custom = json.load(fh)
        unknown = set(custom) - set(DEFAULT_SELECTORS)
        if unknown:
            LOG.warning("Claves de selector desconocidas (se ignoran): %s", ", ".join(unknown))
        selectors.update({k: v for k, v in custom.items() if k in DEFAULT_SELECTORS})
    return selectors


def run_pipeline(cfg: ScraperConfig) -> int:
    stats = RunStats()
    t0 = time.perf_counter()

    records = RealEstateScraper(cfg, stats).run()
    if not records:
        LOG.error("No se extrajo ningún anuncio. Revisa la URL o los selectores CSS.")
        print_summary(stats, pd.DataFrame(), None, cfg, time.perf_counter() - t0)
        return 1

    df_report, df_num = build_dataframe(records, cfg, stats)
    df_summary = build_summary(df_num)
    path = export_excel(df_report, df_summary, cfg.output)
    print_summary(stats, df_num, path, cfg, time.perf_counter() - t0)
    return 0


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(asctime)s | %(levelname)-7s | %(message)s", datefmt="%H:%M:%S")
    try:
        selectors = load_selectors(args.selectors)
    except (OSError, json.JSONDecodeError) as exc:
        LOG.error("No se pudo leer el archivo de selectores: %s", exc)
        return 2

    # Pausas: más cortas en DEMO (servidor local), más conservadoras en sitios reales
    d_min = args.delay_min if args.delay_min is not None else (0.3 if args.demo else 1.5)
    d_max = args.delay_max if args.delay_max is not None else (0.8 if args.demo else 3.5)
    if d_max < d_min:
        d_min, d_max = d_max, d_min

    base = dict(max_pages=args.max_pages, delay_min=d_min, delay_max=d_max,
                fx_rate=args.fx_rate, output=args.output, selectors=selectors)
    try:
        if args.demo:
            with DemoServer(port=args.port) as server:
                cfg = ScraperConfig(start_url=server.base_url, bypass_proxy=True, **base)
                return run_pipeline(cfg)
        return run_pipeline(ScraperConfig(start_url=args.url, **base))
    except KeyboardInterrupt:
        LOG.warning("Ejecución interrumpida por el usuario.")
        return 130


if __name__ == "__main__":
    sys.exit(main())
