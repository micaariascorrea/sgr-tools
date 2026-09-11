"""
Calendario de días hábiles de Argentina.

Incluye feriados nacionales (Ley 27.399), feriados trasladables,
días no laborables (p. ej. Jueves Santo) y puentes turísticos
publicados en https://www.argentina.gob.ar/feriados
"""

from __future__ import annotations

from datetime import date, timedelta
from functools import lru_cache


def _easter_sunday(year: int) -> date:
    """Algoritmo de Meeus/Jones/Butcher (calendario gregoriano)."""
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    el = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * el) // 451
    month = (h + el - 7 * m + 114) // 31
    day = ((h + el - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def _transfer_monday(d: date) -> date:
    """
    Art. 6 Ley 27.399:
    martes/miércoles -> lunes anterior; jueves/viernes -> lunes siguiente.
    """
    wd = d.weekday()  # 0=lun ... 6=dom
    if wd in (1, 2):  # mar, mié
        return d - timedelta(days=wd)
    if wd in (3, 4):  # jue, vie
        return d + timedelta(days=7 - wd)
    return d


# Puentes / días no laborables con fines turísticos (decretos anuales)
_PUENTES: dict[int, list[date]] = {
    2025: [date(2025, 5, 2), date(2025, 8, 15), date(2025, 11, 21)],
    2026: [date(2026, 3, 23), date(2026, 7, 10), date(2026, 12, 7)],
}


@lru_cache(maxsize=32)
def feriados_argentina(year: int) -> frozenset[date]:
    """Conjunto de días no hábiles por feriado / no laborable del año."""
    days: set[date] = set()

    # Inamovibles
    days.add(date(year, 1, 1))  # Año Nuevo
    days.add(date(year, 3, 24))  # Memoria
    days.add(date(year, 4, 2))  # Malvinas
    days.add(date(year, 5, 1))  # Trabajo
    days.add(date(year, 5, 25))  # Revolución de Mayo
    days.add(date(year, 6, 20))  # Belgrano
    days.add(date(year, 7, 9))  # Independencia
    days.add(date(year, 12, 8))  # Inmaculada Concepción
    days.add(date(year, 12, 25))  # Navidad

    easter = _easter_sunday(year)
    carnival_tue = easter - timedelta(days=47)
    carnival_mon = carnival_tue - timedelta(days=1)
    days.add(carnival_mon)
    days.add(carnival_tue)
    days.add(easter - timedelta(days=3))  # Jueves Santo (no laborable)
    days.add(easter - timedelta(days=2))  # Viernes Santo

    # Trasladables (observados)
    for md in ((6, 17), (8, 17), (10, 12), (11, 20)):
        days.add(_transfer_monday(date(year, md[0], md[1])))

    for p in _PUENTES.get(year, []):
        days.add(p)

    return frozenset(days)


def es_dia_habil(d: date) -> bool:
    if d.weekday() >= 5:
        return False
    return d not in feriados_argentina(d.year)


def sumar_dias_habiles(inicio: date, n: int) -> date:
    """Suma n días hábiles a partir del día siguiente a `inicio`."""
    if n < 0:
        raise ValueError("n debe ser >= 0")
    if n == 0:
        return inicio
    cur = inicio
    left = n
    while left > 0:
        cur += timedelta(days=1)
        if es_dia_habil(cur):
            left -= 1
    return cur


def dia_habil_anterior(d: date) -> date:
    """Si `d` no es hábil (fin de semana o feriado), retrocede hasta el hábil previo."""
    cur = d
    while not es_dia_habil(cur):
        cur -= timedelta(days=1)
    return cur


def sumar_dias_corridos(inicio: date, n: int) -> date:
    """Suma n días corridos (calendario, sin excluir fines de semana ni feriados)."""
    if n < 0:
        raise ValueError("n debe ser >= 0")
    return inicio + timedelta(days=n)
