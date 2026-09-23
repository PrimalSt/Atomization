"""Палитра и шрифты презентации, вычисленные из ``ThemeConfig``."""

import string
from dataclasses import dataclass

from pptx.dml.color import RGBColor

from src.schemas import ThemeConfig

WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def hex_to_rgb(value: str) -> RGBColor:
    """«1A365D» или «#1a365d» → RGBColor."""
    normalized = value.strip().removeprefix("#")
    if len(normalized) != 6 or any(char not in string.hexdigits for char in normalized):
        raise ValueError(f"Ожидается цвет в формате RRGGBB, получено «{value}»")
    return RGBColor.from_string(normalized.upper())


def blend(color: RGBColor, other: RGBColor, share: float) -> RGBColor:
    """Смешивает цвета: share — доля ``other`` (0 — исходный цвет, 1 — ``other``)."""
    return RGBColor(*(round(own + (target - own) * share) for own, target in zip(color, other)))


def relative_luminance(color: RGBColor) -> float:
    """Относительная яркость по WCAG 2.x: 0 — чёрный, 1 — белый."""

    def linear(channel: int) -> float:
        value = channel / 255
        return value / 12.92 if value <= 0.04045 else ((value + 0.055) / 1.055) ** 2.4

    red, green, blue = (linear(channel) for channel in color)
    return 0.2126 * red + 0.7152 * green + 0.0722 * blue


def contrast_ratio(first: RGBColor, second: RGBColor) -> float:
    """Контраст по WCAG: от 1 (одинаковые цвета) до 21 (чёрный на белом)."""
    lighter, darker = sorted((relative_luminance(first), relative_luminance(second)), reverse=True)
    return (lighter + 0.05) / (darker + 0.05)


@dataclass(frozen=True, slots=True)
class Theme:
    """Цвета бренда из конфига и нейтральные цвета интерфейса.

    Цвета статусов (лучше/хуже среднего) намеренно не зависят от бренда:
    зелёный и красный должны читаться одинаково в отчётах любого клиента.
    """

    primary: RGBColor
    accent: RGBColor
    background: RGBColor
    heading_font: str
    body_font: str
    surface: RGBColor = WHITE
    text: RGBColor = RGBColor(0x1F, 0x29, 0x37)
    muted: RGBColor = RGBColor(0x5F, 0x6B, 0x7A)
    border: RGBColor = RGBColor(0xE2, 0xE8, 0xF0)
    good: RGBColor = RGBColor(0x1B, 0x7F, 0x3B)
    good_fill: RGBColor = RGBColor(0xE3, 0xF4, 0xE8)
    bad: RGBColor = RGBColor(0xB4, 0x23, 0x18)
    bad_fill: RGBColor = RGBColor(0xFD, 0xE8, 0xE7)

    @classmethod
    def from_config(cls, config: ThemeConfig) -> "Theme":
        return cls(
            primary=hex_to_rgb(config.primary_color),
            accent=hex_to_rgb(config.accent_color),
            background=hex_to_rgb(config.bg_color),
            heading_font=config.heading_font,
            body_font=config.body_font,
        )

    def text_on(self, fill: RGBColor) -> RGBColor:
        """Белый или тёмный текст — тот, что контрастнее на заданной заливке."""
        return max((WHITE, self.text), key=lambda candidate: contrast_ratio(candidate, fill))

    @property
    def primary_soft(self) -> RGBColor:
        """Едва заметный оттенок основного цвета — чередование строк таблицы."""
        return blend(self.primary, WHITE, 0.93)

    @property
    def accent_soft(self) -> RGBColor:
        """Светлый акцент — строка «ИТОГО», нейтральная подсветка."""
        return blend(self.accent, WHITE, 0.8)

    @property
    def accent_on_primary(self) -> RGBColor:
        """Акцент для текста на фоне primary, если он там читается (контраст ≥ 3 для крупного текста)."""
        if contrast_ratio(self.accent, self.primary) >= 3:
            return self.accent
        return self.text_on(self.primary)
