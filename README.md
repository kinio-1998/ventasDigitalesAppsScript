# 📊 Automatización de Reportes de Ventas Quincenales en Google Sheets

Este proyecto en Google Apps Script automatiza el procesamiento de ventas registradas en diferentes hojas mensuales dentro de una hoja de cálculo de Google Sheets. Calcula los totales y cantidades de productos vendidos en la **primera y segunda quincena de cada mes**, y genera automáticamente nuevas hojas mensuales usando una plantilla base.

---

## 🚀 Funcionalidades Principales

- 🔄 Recorre todas las hojas del documento (excepto `ARTICULOS` y `Plantilla`).
- 📆 Identifica la quincena en la que ocurrió cada venta.
- 🧮 Suma los productos vendidos y los totales de venta para cada quincena.
- 🧾 Inserta los resultados directamente en las celdas `M2:M3` (cantidades) y `N2:N3` (totales).
- 🆕 Genera automáticamente una nueva hoja cada día primero del mes basada en una plantilla predefinida.

---

## 🧱 Estructura del Script

| Función        | Descripción |
|----------------|-------------|
| `actualizar()` | Recorre todas las hojas activas, filtra ventas válidas, calcula ventas quincenales y actualiza los resultados en la hoja. |
| `contar(hoja)` | Identifica a qué quincena pertenece cada venta (1–15 o 16–31) y acumula las cantidades y montos. Distingue si la venta fue por Shopify u otro canal. |
| `agregarHoja()`| Duplica la hoja `Plantilla` y la renombra con el nombre del mes actual + año, solo si es día 1. Inicializa las fórmulas y bordes necesarios para la nueva hoja. |

---

## 🧩 Variables Globales

- `datosFiltrados`: Acumulador temporal de datos válidos por hoja.
- `contador1ra`, `contador2da`: Cantidades de productos vendidos por quincena.
- `total1ra`, `total2da`: Monto total de ventas por quincena.
- `meses`: Lista con nombres de los meses en español, usada para comparar fechas de las ventas.
- `totalHojas`: Lista con las hojas válidas para procesar.

---

## ⚙️ Instalación y Uso

1. Abre tu hoja de cálculo de Google.
2. Ve al menú `Extensiones > Apps Script`.
3. Copia y pega el código en el editor (o distribúyelo entre archivos si deseas mantenerlo organizado).
4. Guarda el proyecto.
5. Ejecuta la función `actualizar()` para comenzar el procesamiento.
6. (Opcional) Usa `agregarHoja()` programado con un trigger diario para que el sistema cree nuevas hojas mensuales automáticamente.

---

## 📌 Consideraciones

- La hoja `Plantilla` debe existir y tener el formato base para las nuevas hojas.
- Los datos de ventas deben estar en el rango `A2:J` de cada hoja mensual.
- Las columnas deben tener la siguiente estructura:

| Columna | Contenido |
|---------|-----------|
| A       | SKU o ID de producto |
| B       | Nombre de producto |
| C       | Precio unitario |
| D       | Fecha de venta |
| E       | Cantidad vendida |
| F       | Monto por otro canal |
| G       | Descuento |
| H       | Impuestos |
| I       | Comentarios |
| J       | Canal de venta (ej. `Shopify`) |

- Los datos de calculos deben estar en el rango `L2:N4` es una tabla fija.
- Las columnas deben tener la siguiente estructura:

| Columna | Contenido |
|---------|-----------|
| L       | QUINCENA |
| M       | ARTICULOS VENDIDOS |
| N       | TOTAL |
	 	
---

## ✍️ Autor

Proyecto desarrollado por Carlos Duarte, con el objetivo de automatizar reportes internos en hojas de cálculo de Google para ventas digitales.

---

## 🧾 Licencia

Este proyecto es de uso interno. Puedes modificarlo y adaptarlo según las necesidades de tu empresa o proyecto personal.

