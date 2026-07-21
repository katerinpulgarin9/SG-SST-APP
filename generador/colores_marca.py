# -*- coding: utf-8 -*-
"""Extrae colores de marca suaves a partir del logo de la empresa."""
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


def _lighten(rgb: tuple[int, int, int], factor: float) -> tuple[int, int, int]:
    r, g, b = rgb
    return (
        min(255, int(r + (255 - r) * factor)),
        min(255, int(g + (255 - g) * factor)),
        min(255, int(b + (255 - b) * factor)),
    )


def _desaturate(rgb: tuple[int, int, int], factor: float = 0.35) -> tuple[int, int, int]:
    """Acerca el color al gris para bajar ruido visual."""
    r, g, b = rgb
    gray = int(0.2126 * r + 0.7152 * g + 0.0722 * b)
    return (
        int(r + (gray - r) * factor),
        int(g + (gray - g) * factor),
        int(b + (gray - b) * factor),
    )


def _suavizar(rgb: tuple[int, int, int], lighten: float, desat: float = 0.3) -> tuple[int, int, int]:
    return _lighten(_desaturate(rgb, desat), lighten)


def colores_desde_logo(logo_path: Optional[Path | str]) -> dict[str, str]:
    """Paleta suave derivada del logo: pasteles claros, texto oscuro.

    - primary: acento claro (barras de titulo)
    - secondary: un poco mas claro
    - light: fondo muy suave
    - on_primary: texto oscuro (legible sobre fondos claros)
    """
    # Neutro suave por defecto (sin Forest Green fuerte)
    default = {
        "primary": "B8C9DA",
        "secondary": "D0DCE8",
        "light": "EEF3F7",
        "on_primary": "243447",
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
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[3])
        q = bg.quantize(colors=16, method=Image.Quantize.MEDIANCUT).convert("RGB")
        counts: Counter[tuple[int, int, int]] = Counter()
        for r, g, b in q.getdata():
            if _luma(r, g, b) > 0.92 or _luma(r, g, b) < 0.08:
                continue
            if _sat(r, g, b) < 0.10:
                continue
            key = (r // 8 * 8, g // 8 * 8, b // 8 * 8)
            counts[key] += 1
        if not counts:
            return default

        ranked = sorted(
            counts.items(),
            key=lambda kv: (kv[1] * (0.35 + _sat(*kv[0])),),
            reverse=True,
        )
        base = ranked[0][0]
        base2 = base
        for rgb, _n in ranked[1:]:
            if abs(rgb[0] - base[0]) + abs(rgb[1] - base[1]) + abs(rgb[2] - base[2]) > 50:
                base2 = rgb
                break

        # Tonos claros / pastel (poco ruido)
        primary = _suavizar(base, lighten=0.62, desat=0.28)
        secondary = _suavizar(base2 if base2 != base else base, lighten=0.72, desat=0.32)
        light = _suavizar(base, lighten=0.88, desat=0.40)

        # Evitar que queden todavia oscuros
        if _luma(*primary) < 0.72:
            primary = _lighten(primary, 0.45)
        if _luma(*secondary) < 0.80:
            secondary = _lighten(secondary, 0.40)
        if _luma(*light) < 0.90:
            light = _lighten(light, 0.35)

        return {
            "primary": _hex(primary),
            "secondary": _hex(secondary),
            "light": _hex(light),
            "on_primary": "243447",  # texto oscuro sobre pasteles
        }
    except Exception:
        return default


def es_relleno_marca_plantilla(hex_color: str) -> bool:
    """True si el color parece branding de plantilla (verdes / azules tipicos)."""
    h = (hex_color or "").upper().replace("#", "")
    if len(h) == 8:
        h = h[2:]
    if len(h) != 6:
        return False
    try:
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
    except ValueError:
        return False
    if _sat(r, g, b) < 0.12:
        return False
    if _luma(r, g, b) > 0.93 or _luma(r, g, b) < 0.05:
        return False
    if g >= r and g >= b and _sat(r, g, b) >= 0.15:
        return True
    if b >= r and (b - r) > 20 and _sat(r, g, b) >= 0.15:
        return True
    if _sat(r, g, b) >= 0.35 and _luma(r, g, b) < 0.75:
        return True
    return False


def elegir_tono_marca(hex_color: str, palette: dict[str, str]) -> str:
    """Mapea rellenos de plantilla a tonos claros de la paleta suave."""
    h = (hex_color or "").upper().replace("#", "")
    if len(h) == 8:
        h = h[2:]
    try:
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
    except Exception:
        return palette["light"]
    lu = _luma(r, g, b)
    # Fondos oscuros de plantilla -> primary suave (no oscuro)
    # Fondos medios -> secondary
    # Fondos claros -> light
    if lu < 0.45:
        return palette["primary"]
    if lu < 0.75:
        return palette["secondary"]
    return palette["light"]
