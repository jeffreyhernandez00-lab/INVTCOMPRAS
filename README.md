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
- Costo unitario
- Inventario actual

La app incluye un botón para descargar una plantilla con la estructura esperada.
También tolera diferencias de mayúsculas, espacios y acentos en los nombres de columnas.

## Fórmulas de inventario

- Inventario mínimo = Promedio diario * Buffer seguridad
- Punto de reorden = Promedio diario * (Entrega proveedor + Buffer seguridad)
- Inventario máximo = Promedio diario * Días inventario máximo
