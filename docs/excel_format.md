# Formato de los Excel

Todas las posiciones se configuran en `app/backend/config/constants.py`.

## Informe de Economato (entrada)

Informe `.xlsx` exportado del almacén central. La hoja se detecta
automáticamente (se elige la que más filas útiles contiene), por lo que no
importan su nombre ni su posición. La fila 1 es la cabecera.

| Columna | Dato | Notas |
|---|---|---|
| A | Fecha | Fecha de Excel o texto en varios formatos habituales. |
| F | Punto de venta | Con prefijo `XAN - `, que se ignora. |
| G | Grupo de producto | Solo `BEBIDAS`, `ENVASES`, `DROGUERÍA` y `ALIMENTOS/COMIDA`. |
| J | Código de producto | Solo dígitos. |
| L | Nombre del producto | |
| M | Formato | Unidad, caja, etc. |
| P | Cantidad | Admite coma decimal, separador de miles y símbolo €. Las filas con cantidad 0 se ignoran. |
| Q | Precio | |

### Puntos de venta reconocidos

| En el informe | Excel de destino |
|---|---|
| BAR PISCINA | `Bar_Piscina` |
| BAR SALÓN | `Bar_Salon` |
| COMEDOR | `Comedor` |
| MANOLETE, TRACTORIA | `Manolete` |
| STARCAFÉ | `Starcafe` |
| GENERALES BAR Y COMEDOR | `Generales` |

Las filas de otros puntos de venta o grupos se ignoran y se informan en el
resumen de la sincronización.

## Excel mensual (salida)

Se crea a partir de la plantilla del punto de venta la primera vez que llega
una entrega de ese mes:

```
Excel mensuales/2026/JULIO/Bar_Piscina_Julio_2026.xlsx
```

Hoja `extraccion`:

| Posición | Contenido |
|---|---|
| Fila 6 | Cabecera con los días del mes (1–31). |
| Fila 7 en adelante | Un producto por fila, hasta la fila de totales. |
| A | Código |
| B | Nombre |
| C | Stock inicial |
| D | Formato |
| E | Precio |
| G … AK | Cantidad recibida cada día (G = día 1). |
| AL | Total extraído en el mes (fórmula). |
| AM | Stock total (fórmula). |
| AN | Valor (fórmula). |

Cuando llega un producto que no existe, se inserta una fila nueva justo antes
de la fila de totales copiando el formato y las fórmulas de la fila anterior,
y se amplían las fórmulas de totales. El mismo producto se añade también a la
plantilla para los meses siguientes.

La hoja oculta `__SYNC_STATE__` registra las entregas aplicadas (ver
[arquitectura](architecture.md#garantías-de-integridad)). No debe modificarse
a mano.
