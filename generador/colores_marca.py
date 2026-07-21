# -*- coding: utf-8 -*-
"""Extrae colores de marca a partir del logo de la empresa."""
from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Optional


def _luma(r: int, g: int, b: int) -> float:
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0


def _sat(r: int, g: int, b: int) -> float:
    mx = max(r, g, b)
    mn = min(r, g, b)
    if mx == 0:
        return 0.0
    return (mx - mn) / mx


def _hex(rgb: tuple[int, int, int]) -> str:
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def _lighten(rgb: tuple[int, int, int], factor: float = 0.75) -> tuple[int, int, int]:
    r, g, b = rgb
    return (
        min(255, int(r + (255 - r) * factor)),
        min(255, int(g + (255 - g) * factor)),
        min(255, int(b + (255 - b) * factor)),
    )


def _darken(rgb: tuple[int, int, int], factor: float = 0.25) -> tuple[int, int, int]:
    r, g, b = rgb
    return (
        max(0, int(r * (1 - factor))),
        max(0, int(g * (1 - factor))),
        max(0, int(b * (1 - factor))),
    )


def colores_desde_logo(logo_path: Optional[Path | str]) -> dict[str, str]:
    """Devuelve paleta {primary, secondary, light, on_primary} en hex RRGGBB.

    Si no hay logo o falla, usa una paleta neutra profesional (no Forest Green).
    """
    default = {
        "primary": "1F4E79",
        "secondary": "2E75B6",
        "light": "D6E3F0",
        "on_primary": "FFFFFF",
    }
    if not logo_path:
        return default
    path = Path(logo_path)
    if not path.exists():
        return default
    try:
        from PIL import Image
    except Exception:
        return default

    try:
        img = Image.open(path).convert("RGBA")
        img.thumbnail((160, 160))
        # Fondo blanco para transparencias
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[3])
        # Cuantizar para agrupar tonos
        q = bg.quantize(colors=16, method=Image.Quantize.MEDIANCUT).convert("RGB")
        pixels = list(q.getdata())
        counts: Counter[tuple[int, int, int]] = Counter()
        for r, g, b in pixels:
            # Ignorar casi blanco / casi negro
            if _luma(r, g, b) > 0.92 or _luma(r, g, b) < 0.08:
                continue
            if _sat(r, g, b) < 0.12 and _luma(r, g, b) > 0.75:
                continue
            # Agrupar un poco
            key = (r // 8 * 8, g // 8 * 8, b // 8 * 8)
            counts[key] += 1
        if not counts:
            return default

        # Preferir colores con saturacion decente
        ranked = sorted(
            counts.items(),
            key=lambda kv: (kv[1] * (0.4 + _sat(*kv[0])), -abs(_luma(*kv[0]) - 0.4)),
            reverse=True,
        )
        primary = ranked[0][0]
        # Secundario: siguiente tono distinto
        secondary = primary
        for rgb, _n in ranked[1:]:
            if abs(rgb[0] - primary[0]) + abs(rgb[1] - primary[1]) + abs(rgb[2] - primary[2]) > 60:
                secondary = rgb
                break
        if secondary == primary:
            secondary = _lighten(primary, 0.35)
        light = _lighten(primary, 0.82)
        on_primary = "FFFFFF" if _luma(*primary) < 0.55 else "1A1A1A"
        # Asegurar primary no demasiado claro para barras de titulo
        if _luma(*primary) > 0.65:
            primary = _darken(primary, 0.35)
        return {
            "primary": _hex(primary),
            "secondary": _hex(secondary),
            "light": _hex(light),
            "on_primary": on_primary,
        }
    except Exception:
        return default


def es_relleno_marca_plantilla(hex_color: str) -> bool:
    """True si el color parece branding de plantilla (verdes Forest / azules tipicos)."""
    h = (hex_color or "").upper().replace("#", "")
    if len(h) == 8:  # AARRGGBB
        h = h[2:]
    if len(h) != 6:
        return False
    try:
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
    except ValueError:
        return False
    # Blancos / grises / negros: no tocar
    if _sat(r, g, b) < 0.12:
        return False
    if _luma(r, g, b) > 0.93 or _luma(r, g, b) < 0.05:
        return False
    # Verdes / azules-verdes tipicos de plantillas SST
    if g >= r and g >= b and _sat(r, g, b) >= 0.15:
        return True
    if b >= r and (b - r) > 20 and _sat(r, g, b) >= 0.15:
        return True
    # Cualquier relleno bastante saturado en encabezado tambien se recolorea
    if _sat(r, g, b) >= 0.35 and _luma(r, g, b) < 0.75:
        return True
    return False


def elegir_tono_marca(hex_color: str, palette: dict[str, str]) -> str:
    """Mapea un relleno de plantilla a primary/secondary/light segun luminosidad."""
    h = (hex_color or "").upper().replace("#", "")
    if len(h) == 8:
        h = h[2:]
    try:
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
    except Exception:
        return palette["primary"]
    lu = _luma(r, g, b)
    if lu >= 0.75:
        return palette["light"]
    if lu >= 0.45:
        return palette["secondary"]
    return palette["primary"]
