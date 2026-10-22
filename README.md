# Verificación de Solicitudes de Traspaso

App en Streamlit para validar solicitudes de traspaso entre locales antes de grabarlas, de modo que no se
deje a un local sin su inventario de exhibición, y para proponer desde qué locales sacar la mercadería.

## Qué hace

Para cada solicitud (origen → destino, estadístico, cantidad) cubre la cantidad en este orden:

1. **Origen solicitado sin bajar de exhibición**: lo que el origen puede ceder quedando en su exhibición.
2. **Alternativos con sobrante**: otros locales con físico mayor a su exhibición, de mayor a menor sobrante,
   **sin límite de locales**. Nunca quedan bajo exhibición.
3. **Origen bajo exhibición**: si aún falta, el origen pone lo que tenga, aunque quede bajo exhibición.
4. **Extremos**: si el origen ya está en 0, se toma de otros locales aunque queden bajo exhibición, de mayor a
   menor stock físico.

Lo que no se puede cubrir queda como **pendiente**. El stock se descuenta en tiempo real: lo que un local
cede en una fila ya no está disponible para las siguientes (las primeras filas tienen prioridad).

Reglas: los alternativos nunca son el local destino ni (por defecto) un local que ya es origen del mismo
estadístico en la solicitud. Opcionalmente se pueden restringir al mismo supervisor del origen.

## Archivos de entrada

| Archivo | Columnas necesarias |
|---|---|
| Base del BI | `# Local`, `ESTADISTICO`, `Físico`, `Uni Exhibición` (opcionales: `DESCRIPCION`, `SUPERV.`, `DESIGNACION`, `CONSUMOS ACU`) |
| Solicitudes | `Loc. Origen`, `Loc. Destino`, `Estadístico`, `Cantidad` (acepta el reporte de Solicitud de Traspasos y la hoja TRASLADOS SOLICITADOS del Excel anterior) |

Los encabezados se detectan solos aunque no estén en la primera fila. Los códigos pueden venir como `189` o
`189-GUASMO LAS ESCLUSAS`, y `51404000` o `051404000_MEDIAS…`. Las solicitudes también se pueden pegar
directamente en la app.

## Resultado (Excel descargable)

- **RESUMEN**: unidades por paso, pendientes, validación inicial.
- **TRASPASOS A GRABAR**: lista consolidada; columnas A:D listas para grabar en el sistema.
- **REPORTE BAJO EXHIBICIÓN**: cada local + ítem que queda bajo exhibición, con destinos y motivo.
- **DETALLE POR SOLICITUD**: cómo se cubrió cada fila.
- **PENDIENTES**: solo si hay unidades sin cubrir.

## Ejecutar en tu computadora

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicar en Streamlit Community Cloud

1. Sube esta carpeta a un repositorio de GitHub (los archivos `.xlsx`/`.csv` de datos están excluidos en `.gitignore`).
2. Entra a https://share.streamlit.io, inicia sesión con GitHub y elige **Create app**.
3. Selecciona el repositorio, la rama `main` y el archivo principal `app.py`. Clic en **Deploy**.

Los archivos que subes a la app se procesan en memoria y no se guardan en el repositorio.

## Pruebas

```bash
pip install pytest
python -m pytest
```

## Estructura

```
app.py          interfaz Streamlit
motor.py        lectura de archivos y lógica de asignación
exportar.py     Excel de resultados con formato
tests/          pruebas con datos de ejemplo
```
