# Sugerido de compra automático

Aplicación web sencilla en Streamlit para generar un sugerido de compra desde un archivo Excel con ventas, costo e inventario actual.

## Ejecutar

```powershell
pip install -r requirements.txt
streamlit run app.py
```

## Publicar en Streamlit Community Cloud

1. Crea un repositorio en GitHub.
2. Sube estos archivos del proyecto:
   - `streamlit_app.py`
   - `app.py`
   - `requirements.txt`
   - `.streamlit/config.toml`
3. Entra a `https://share.streamlit.io`.
4. Selecciona tu repositorio, la rama principal y el archivo `streamlit_app.py`.
5. Presiona `Deploy`.

Cuando termine, Streamlit te dará un enlace similar a:

```text
https://nombre-de-tu-app.streamlit.app
```

Ese enlace lo podrán abrir otras computadoras desde cualquier navegador.

## Columnas obligatorias del Excel

- Código de producto
- Nombre de producto
- Laboratorio
- Venta mes 1
- Venta mes 2
- Venta mes 3
- Venta mes 4
- Venta mes 5
- Venta mes 6
- Costo unitario
- Inventario actual

La app incluye un botón para descargar una plantilla con la estructura esperada.
La plantilla incluye una hoja `Productos`, una hoja `Responsables laboratorio`, una hoja `Parámetros ABC` y una hoja `Configuración`.
También tolera diferencias de mayúsculas, espacios y acentos en los nombres de columnas.

La hoja `Responsables laboratorio` permite definir una persona responsable por cada laboratorio.
La app también muestra un apartado editable para revisar o completar esos responsables antes de calcular.

## Fórmulas de inventario

- Inventario mínimo = Promedio diario * Buffer seguridad
- Punto de reorden = Promedio diario * (Entrega proveedor + Buffer seguridad)
- Inventario máximo = Promedio diario * (Entrega proveedor + Buffer seguridad + Días de inventario)
- Diferencia vs máximo = Inventario máximo - Inventario actual
- Promedio mensual = Consumo 6 meses / 6
- Promedio diario = Promedio mensual / Días hábiles de consumo por mes

Si `Diferencia vs máximo` es negativa, el producto aparece como `Exceso`.

## Formato de resultados

Solo `Costo unitario` y `Total estimado compra` se muestran con decimales.
Los promedios, inventarios, punto de reorden, sugeridos, porcentajes y demás cantidades se muestran como números enteros.
