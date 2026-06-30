# Gastos Bleif Meals — Apps Script Automation

Conjunto de scripts de Google Apps Script para automatizar la gestión mensual del archivo de gastos **"Gastos Bleif Meals"** y su gestor de recetas, todo dentro de un mismo Google Sheets.

## ¿Qué hace?

El proyecto resuelve dos flujos de trabajo distintos sobre la misma planilla:

### 1. Apertura del archivo del próximo mes
Cada mes hay que duplicar el archivo completo de gastos, dejarlo guardado en la carpeta correspondiente de Drive y prepararlo con los saldos iniciales del mes que arranca. En vez de hacerlo a mano, el menú **Scripts → Crear Archivo Mes Siguiente** ejecuta todo el proceso de punta a punta.

### 2. Gestión del catálogo de recetas
La planilla también funciona como base de recetas. Hay funciones para crear nuevas recetas a partir de una plantilla protegida, mantener actualizado el listado general y mantener las pestañas ordenadas alfabéticamente.

## Estructura del proyecto

| Archivo | Responsabilidad |
|---|---|
| `appsscript.gs` | Manifest del proyecto (timezone, runtime V8) |
| `triggers.gs` | Define el menú personalizado `Scripts` en la barra de Sheets |
| `createCopy.gs` | Lógica de duplicación mensual del workbook completo |
| `Code.gs` | Gestión de recetas: creación, protección de columnas y ordenamiento de pestañas |

## Flujo: Crear Archivo Mes Siguiente

Disparado desde el menú `Scripts → Crear Archivo Mes Siguiente` (función `duplicarWorkbookYLimpiar`):

1. **Detecta el próximo mes** a partir del nombre del archivo actual (ej. `"Gastos Bleif Meals - JULIO 2026"` → calcula `AGOSTO 2026`), con manejo automático de cambio de año si el mes actual es diciembre.
2. **Confirma con el usuario** el nombre del nuevo archivo antes de crear nada.
3. **Ubica la carpeta de destino** en Drive siguiendo la estructura `Raíz / Año / Mes` (ej. `2026 / 08 - Agosto`), creando las carpetas que falten.
4. **Duplica el archivo completo** (todas las pestañas, formato y fórmulas) en la carpeta del nuevo mes.
5. **Limpia los rangos con datos del mes anterior** en la copia (`RANGOS_A_BORRAR`), configurable por hoja.
6. **Traspasa saldos finales → iniciales**: copia como *valores* (sin fórmulas) los totales finales del mes que cierra hacia las celdas iniciales del mes que arranca (`TRASPASO_VALORES`), incluyendo soporte para rangos de columna abierta (ej. `M4:M`) calculando automáticamente el largo real de la tabla.

### Configuración

Toda la lógica de negocio vive en dos objetos al inicio de `createCopy.gs`:

```javascript
// Qué rangos limpiar en la copia, por hoja
const RANGOS_A_BORRAR = {
  'Ventas': ['A3:B', 'E3:G', 'L3:BY'],
  'Produccion': ['F8:F12', 'L8:BW12'],
  'Stock y Compras': ['G4:M'],
};

// Qué valores finales del mes anterior pasan a ser los iniciales del nuevo mes
const TRASPASO_VALORES = [
  { hoja: 'Produccion', origen: 'H15:BW15', destino: 'H7:BW7' },
  { hoja: 'Stock y Compras', origen: 'M4:M', destino: 'G4:G' },
];
```

## Flujo: Gestión de recetas

- **`duplicarPestanaConNombre`** — Crea una nueva receta a partir de la pestaña `Recetas Template`, pide el nombre por prompt, evita duplicados, protege las columnas clave (`A`, `C`, `E`) con advertencia y actualiza el listado general en la pestaña `Recetas ---->`.
- **`ordenarPestanasPorNombre`** — Ordena alfabéticamente (es-AR) todas las pestañas ubicadas después de `Recetas ---->`, sin tocar el resto de la estructura del libro.

## Requisitos

- El archivo debe estar guardado en Drive siguiendo la estructura `.../Año/Mes/archivo`.
- El nombre del archivo debe mantener el patrón `"Gastos Bleif Meals - <MES> <AÑO>"` para que la detección automática del próximo mes funcione.
- Permisos de Drive y Sheets (Apps Script los solicita la primera vez que se ejecuta cualquier función).

## Instalación

1. Abrí el Google Sheets → **Extensiones → Apps Script**.
2. Copiá cada archivo (`appsscript.gs`, `triggers.gs`, `createCopy.gs`, `Code.gs`) tal cual a su archivo correspondiente en el editor.
3. Guardá y recargá el Sheets — debería aparecer el menú **Scripts** en la barra superior.

## Notas

- Todo el traspaso de valores se hace por `setValues()`, nunca se copian fórmulas: el archivo nuevo arranca "limpio" con valores fijos como saldo inicial.
- No hay ambiente de testing separado — cualquier prueba corre directo sobre el archivo real, así que conviene revisar la configuración antes de ejecutar en producción.
