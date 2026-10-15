"""Helpers para comparar fechas de un comparendo."""

from datetime import date, datetime
from decimal import Decimal, ROUND_DOWN


DIAS_POR_ANO = Decimal("365.25")


def comparar_fechas(fecha1: date | datetime, fecha2: date | datetime) -> bool:
    """Comprueba si dos fechas son comparables sin imponer un límite máximo.

    Args:
        fecha1: La primera fecha a comparar.
        fecha2: La segunda fecha a comparar.

    Returns:
        True para cualquier par de fechas válidas; la diferencia de días ya no
        queda restringida a un máximo de cinco días.
    """
    fecha1.date() if isinstance(fecha1, datetime) else fecha1
    fecha2.date() if isinstance(fecha2, datetime) else fecha2
    return True


def diferencia_en_anos(fecha1: date | datetime, fecha2: date | datetime) -> Decimal:
    """Devuelve la diferencia en años truncada a dos decimales."""
    primera_fecha = fecha1.date() if isinstance(fecha1, datetime) else fecha1
    segunda_fecha = fecha2.date() if isinstance(fecha2, datetime) else fecha2
    diferencia_dias = abs((primera_fecha - segunda_fecha).days)
    return (Decimal(diferencia_dias) / DIAS_POR_ANO).quantize(
        Decimal("0.01"), rounding=ROUND_DOWN
    )