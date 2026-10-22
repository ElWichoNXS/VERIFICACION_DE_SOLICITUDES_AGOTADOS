from pathlib import Path

import motor

D = Path(__file__).parent / "datos"


def _calc(**kw):
    base = motor.leer_base(D / "base_ejemplo.csv")
    sol = motor.leer_solicitudes(D / "solicitudes_ejemplo.csv")
    return motor.calcular(base, sol, motor.Opciones(**kw))


def test_codigos():
    assert motor.codigo("189-GUASMO") == 189
    assert motor.codigo("051404000_MEDIAS") == 51404000
    assert motor.codigo(189.0) == 189


def test_extremos_cuando_origen_sin_stock():
    res = _calc()
    d = res.detalle.set_index("Fila")
    # fila 1: origen 310 en 0 -> sale de otros locales aunque queden bajo exhibición
    assert d.loc[1, "Und. Extremos"] == 12 and d.loc[1, "Pendiente sin Cubrir"] == 0
    # el stock se descuenta entre filas: 20 + 19 + 6 = 45 disponibles en total
    assert res.resumen["paso4"] == 45
    assert d.loc[3, "Pendiente sin Cubrir"] == 17


def test_mas_de_tres_alternativos_sin_bajar_exhibicion():
    res = _calc()
    d = res.detalle.set_index("Fila")
    # ítem 111: origen cede 6, luego 105 (19), 356 (20) y 400-403 (1 c/u) -> 49 sin dejar a nadie bajo exhibición
    assert d.loc[4, "1. Origen sin bajar exhibición"] == 6
    assert d.loc[4, "Und. Alternativos"] == 43
    assert d.loc[4, "Pendiente sin Cubrir"] == 0
    tg = res.traspasos[res.traspasos["Estadístico"] == 111]
    assert (tg["¿Queda Bajo Exhibición?"] == "NO").all()
    assert len(tg) == 7


def test_sin_extremos():
    res = _calc(usar_extremos=False)
    assert res.resumen["paso4"] == 0
    assert res.resumen["pendiente"] == 62
