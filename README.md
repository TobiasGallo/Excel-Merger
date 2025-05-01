# Excel Merger v0.1.0

Una herramienta con interfaz gráfica desarrollada en Python para automatizar la integración de datos entre archivos Excel (.xlsx/.xls) usando coincidencias de claves primarias.

## ✨ Descripción general

**Excel Merger** permite agregar automáticamente datos de un archivo Excel a otro, buscando coincidencias entre columnas clave (como nombres o IDs), conservando los formatos originales (fechas, números, estilos) y generando un nuevo archivo combinado, todo desde una interfaz gráfica sencilla.

## 🧩 Problema que resuelve

En entornos donde no existen claves únicas compartidas entre sistemas, se vuelve tedioso y propenso a errores el cruce manual de datos en Excel. Esta herramienta automatiza ese proceso, aumentando la precisión y reduciendo drásticamente el tiempo de trabajo.

## ⚙️ Funcionalidades técnicas

- Búsqueda inteligente por coincidencia exacta de claves entre archivos.
- Soporte para múltiples apariciones de una misma clave.
- Conservación de formatos complejos (fechas, números, fórmulas).
- Interfaz gráfica con botones para selección de archivos, mezcla y guardado.
- Compatibilidad con `.xlsx` y `.xls`.

## 🖥️ Interfaz gráfica

Desarrollada con Tkinter, incluye:

- Diálogos para elegir archivos.
- Botón "Mezclar datos" para ejecutar la fusión.
- Botón "Guardar archivo" con diálogo para exportar el resultado.
- Mensajes de error y éxito claros.

## ✅ Casos de uso clave

- Agregar datos de clientes a facturas (Facturación).
- Unir SKUs con descripciones de producto (Inventario).
- Consolidar reportes interdepartamentales.
- Vincular tareas y responsables en proyectos.
- Enriquecer tickets de soporte con datos del cliente.

## 🚀 Ventajas competitivas

- ⏱ Ahorro de tiempo (de horas a segundos).
- 🎯 Mayor precisión, sin errores de copiado.
- 🖥 Compatible con Windows, macOS y Linux.

## 🔧 Herramientas utilizadas

- Python 3.7+
- `pandas`
- `openpyxl`
- `tkinter`

## 🔄 Ejemplo de flujo de trabajo

**Archivo Principal:**

| Cliente    | Producto | Fecha       |
|------------|----------|-------------|
| Empresa A  | Lápices  | 2023-10-01  |

**Archivo Secundario:**

| Nombre     | CUIT           |
|------------|----------------|
| Empresa A  | 30-12345678-9  |

**Archivo Combinado:**

| Cliente    | Producto | Fecha       | CUIT           |
|------------|----------|-------------|----------------|
| Empresa A  | Lápices  | 2023-10-01  | 30-12345678-9  |

## ⚠️ Limitaciones

- Solo procesa la primera hoja de cada archivo.
- No admite formatos CSV.
- No valida formatos de campos (ej. CUIT).
- Sin indicadores visuales de progreso.
- Rendimiento limitado con archivos grandes (>50,000 filas).
- Para cambiar columnas clave se requiere editar el código.

## 📁 Repositorio

[https://github.com/TobiasGallo/Excel-Merger](https://github.com/TobiasGallo/Excel-Merger)

## 👤 Desarrollador

**Tobías Gallo**  
📧 tobiasgallo89@gmail.com

---

> Versión: v0.1.0 - Lanzamiento inicial
