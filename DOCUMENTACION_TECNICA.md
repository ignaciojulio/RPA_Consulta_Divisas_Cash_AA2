# 📄 DOCUMENTACIÓN TÉCNICA — RPA Consulta Divisas Cash
> **Versión:** 1.1.0
> **Fecha de actualización:** 09/05/2026
> **Autor:** Ignacio Javier Julio Posada
> **Herramienta:** UiPath Studio — Modern Design (Windows Framework)
> **Estado del proyecto:** ✅ Pipeline completo funcional — Extracción web en modo placeholder

---

## 📑 Tabla de Contenidos

1. [Descripción General del Proyecto](#1-descripción-general-del-proyecto)
2. [Arquitectura Modular](#2-arquitectura-modular)
3. [Flujo de Datos entre Módulos](#3-flujo-de-datos-entre-módulos)
4. [Descripción Técnica por Flujo de Trabajo](#4-descripción-técnica-por-flujo-de-trabajo)
   - 4.1 [Main.xaml — Orquestador Maestro](#41-mainxaml--orquestador-maestro)
   - 4.2 [Inicializacion.xaml — Preparación del Entorno](#42-inicializacionxaml--preparación-del-entorno)
   - 4.3 [Extraer_TRM_USD.xaml — Extracción Dólar](#43-extraer_trm_usdxaml--extracción-dólar)
   - 4.4 [Extraer_TRM_EUR.xaml — Extracción Euro](#44-extraer_trm_eurxaml--extracción-euro)
   - 4.5 [Limpieza_Validacion_Datos.xaml — Transformación y Validación](#45-limpieza_validacion_datosxaml--transformación-y-validación)
   - 4.6 [Registro_Excel.xaml — Persistencia Histórica](#46-registro_excelxaml--persistencia-histórica)
5. [Trazabilidad Completa de Logs](#5-trazabilidad-completa-de-logs)
6. [Requerimientos Funcionales Implementados](#6-requerimientos-funcionales-implementados)
7. [Dependencias y Paquetes NuGet](#7-dependencias-y-paquetes-nuget)
8. [Checklist de Pre-Ejecución](#8-checklist-de-pre-ejecución)
9. [Mantenimiento y Consideraciones de Producción](#9-mantenimiento-y-consideraciones-de-producción)
10. [Historial de Cambios](#10-historial-de-cambios)

---

## 1. Descripción General del Proyecto

| Campo | Detalle |
|---|---|
| **Nombre del Proyecto** | `RPA_Consulta_Divisas_Cash_AA2` |
| **Empresa cliente** | Cash (empresa de divisas) |
| **Objetivo** | Extracción autónoma diaria de la TRM para USD y EUR desde Google Finance, con validación de integridad y registro histórico en Excel |
| **Tipo de proceso** | Process (UiPath Orchestrator compatible) |
| **Framework** | Windows — Modern Design |
| **ID del Proyecto** | `a4f272eb-8d9b-47d4-851b-abbddcd21cc7` |
| **Orchestrator Folder ID** | `7812985` |
| **Número de módulos** | 6 flujos de trabajo `.xaml` |
| **Versión actual** | 1.1.0 |

### Estado del Pipeline por Módulo

| Módulo | Estado | Detalle |
|---|---|---|
| `Main.xaml` | ✅ Funcional | Argumentos completamente cableados (corregido 09/05/2026) |
| `Inicializacion.xaml` | ✅ Funcional | `PathExcel` expuesto como `OutArgument` |
| `Extraer_TRM_USD.xaml` | ⚠️ Placeholder | Retorna `"5000.00"` fijo — UI Automation pendiente |
| `Extraer_TRM_EUR.xaml` | ⚠️ Placeholder | Retorna `"4900.00"` fijo — UI Automation pendiente |
| `Limpieza_Validacion_Datos.xaml` | ✅ Funcional | Validación RF-02 operativa |
| `Registro_Excel.xaml` | ✅ Funcional | Migrado a `WriteRangeX` + `ExcelApplicationCard` |

### Alcance del Proceso

El robot realiza de forma completamente autónoma las siguientes tareas:
1. Limpieza del entorno de trabajo (cierre de instancias Excel abiertas).
2. Verificación y auto-creación del archivo maestro Excel si no existe.
3. *(Pendiente — RF-01)* Navegación web a Google Finance en modo InPrivate y extracción del precio actual de USD/COP y EUR/COP.
4. Limpieza, conversión numérica y validación de integridad de los datos (RF-02).
5. Registro histórico de los datos en el archivo `Control_Divisas.xlsx` sin sobrescribir el historial (RF-04).

---

## 2. Arquitectura Modular

```
RPA_Consulta_Divisas_Cash_AA2/
│
├── Main.xaml                        ← Orquestador maestro (entry point)
├── Inicializacion.xaml              ← Módulo 1: Preparación del entorno
├── Extraer_TRM_USD.xaml             ← Módulo 2: Extracción TRM Dólar  [⚠️ placeholder]
├── Extraer_TRM_EUR.xaml             ← Módulo 3: Extracción TRM Euro   [⚠️ placeholder]
├── Limpieza_Validacion_Datos.xaml   ← Módulo 4: Transformación y validación
├── Registro_Excel.xaml              ← Módulo 5: Persistencia en Excel
│
├── Control_Divisas.xlsx             ← Archivo de datos históricos (auto-creado)
├── project.json                     ← Configuración del proyecto UiPath
├── README.md                        ← Descripción general del proyecto
└── DOCUMENTACION_TECNICA.md         ← Este documento
```

### Principios de Diseño Aplicados

| Principio | Implementación |
|---|---|
| **Modularidad** | Cada responsabilidad encapsulada en su propio `.xaml` |
| **Portabilidad** | `Path.Combine(Environment.CurrentDirectory, ...)` para rutas relativas |
| **Alta Disponibilidad** | `Retry Scope` con 3 intentos en cada extracción web |
| **Tolerancia a Fallos** | `Try-Catch` en extracción web y escritura Excel |
| **Auto-Bootstrap** | El robot crea su propio Excel si no lo encuentra |
| **Sesión Limpia** | `Use Application/Browser` en modo InPrivate (sin cookies) |
| **Solo Actividades Modern** | Cero actividades Classic de Excel — compatible con `UiPath.Excel.Activities 3.5.1+` |

---

## 3. Flujo de Datos entre Módulos

```
┌──────────────────────────────────────────────────────────────┐
│                         MAIN.XAML                            │
│                      (Try-Catch Global)                      │
│  Variables: PathExcel, strValorUSD, strValorEUR,             │
│             numValorUSD, numValorEUR                         │
└───────────────────────┬──────────────────────────────────────┘
                        │
           ┌────────────▼────────────┐
           │    Inicializacion.xaml  │
           │  OUT: PathExcel (Str)   │
           └────────────┬────────────┘
                        │ [PathExcel]
           ┌────────────▼────────────────┐
           │   Extraer_TRM_USD.xaml      │
           │   OUT: strValorUSD (Str)    │  ⚠️ Placeholder: "5000.00"
           └────────────┬────────────────┘
                        │ [strValorUSD]
           ┌────────────▼────────────────┐
           │   Extraer_TRM_EUR.xaml      │
           │   OUT: strValorEUR (Str)    │  ⚠️ Placeholder: "4900.00"
           └────────────┬────────────────┘
                        │ [strValorUSD] + [strValorEUR]
      ┌─────────────────▼──────────────────────┐
      │      Limpieza_Validacion_Datos.xaml     │
      │  IN:  strValorUSD, strValorEUR          │
      │  OUT: numValorUSD (Double)              │
      │       numValorEUR (Double)              │
      └─────────────────┬──────────────────────┘
                        │ [numValorUSD] + [numValorEUR] + [PathExcel]
           ┌────────────▼────────────┐
           │   Registro_Excel.xaml   │
           │  IN: in_numValorUSD     │
           │      in_numValorEUR     │
           │      in_PathExcel       │
           └─────────────────────────┘
```

> **Corrección v1.1.0 — Root Cause Fix:** Todos los `InvokeWorkflowFile` en `Main.xaml` tenían diccionarios de argumentos **vacíos** desde la versión inicial. Esto causaba que los valores extraídos nunca llegaran a los módulos de validación y registro (`strValorUSD = ""`, `strValorEUR = ""`). Corregido el 09/05/2026 con el cableado completo de `InArgument` y `OutArgument` en los 5 invocadores.

---

## 4. Descripción Técnica por Flujo de Trabajo

---

### 4.1 `Main.xaml` — Orquestador Maestro

**Tipo:** Entry Point | **Argumentos propios:** Ninguno

#### Variables del Flujo Principal

| Variable | Tipo | Propósito |
|---|---|---|
| `PathExcel` | `String` | Ruta al archivo Excel (recibida de `Inicializacion`) |
| `strValorUSD` | `String` | Texto crudo del precio USD (recibido de `Extraer_TRM_USD`) |
| `strValorEUR` | `String` | Texto crudo del precio EUR (recibido de `Extraer_TRM_EUR`) |
| `numValorUSD` | `Double` | Valor numérico validado del Dólar (recibido de `Limpieza_Validacion`) |
| `numValorEUR` | `Double` | Valor numérico validado del Euro (recibido de `Limpieza_Validacion`) |

#### Tabla de Cableado de Argumentos — InvokeWorkflowFile (v1.1.0)

| Flujo Invocado | Argumento | Dirección | Expresión en Main |
|---|---|---|---|
| `Inicializacion.xaml` | `PathExcel` | `OutArgument<String>` | `[PathExcel]` |
| `Extraer_TRM_USD.xaml` | `strValorUSD` | `OutArgument<String>` | `[strValorUSD]` |
| `Extraer_TRM_EUR.xaml` | `strValorEUR` | `OutArgument<String>` | `[strValorEUR]` |
| `Limpieza_Validacion_Datos.xaml` | `strValorUSD` | `InArgument<String>` | `[strValorUSD]` |
| `Limpieza_Validacion_Datos.xaml` | `strValorEUR` | `InArgument<String>` | `[strValorEUR]` |
| `Limpieza_Validacion_Datos.xaml` | `numValorUSD` | `OutArgument<Double>` | `[numValorUSD]` |
| `Limpieza_Validacion_Datos.xaml` | `numValorEUR` | `OutArgument<Double>` | `[numValorEUR]` |
| `Registro_Excel.xaml` | `in_numValorUSD` | `InArgument<Double>` | `[numValorUSD]` |
| `Registro_Excel.xaml` | `in_numValorEUR` | `InArgument<Double>` | `[numValorEUR]` |
| `Registro_Excel.xaml` | `in_PathExcel` | `InArgument<String>` | `[PathExcel]` |

#### Estructura de Control

```
Sequence
 ├── ManualTrigger
 ├── Log Message [Info]
 └── Try-Catch (System.Exception)
      ├── Try:
      │    ├── Invoke → Inicializacion.xaml             [OUT: PathExcel]
      │    ├── Invoke → Extraer_TRM_USD.xaml             [OUT: strValorUSD]
      │    ├── Invoke → Extraer_TRM_EUR.xaml             [OUT: strValorEUR]
      │    ├── Invoke → Limpieza_Validacion_Datos.xaml   [IN: strs / OUT: nums]
      │    ├── Invoke → Registro_Excel.xaml              [IN: nums + PathExcel]
      │    └── Log Message [Info]
      └── Catch (System.Exception):
           └── (vacío — pendiente agregar rethrow y log de error)
```

> ⚠️ **Mejora pendiente:** Agregar `Log Message [Error]` y `Rethrow` en el bloque `Catch` para propagar errores críticos al Orchestrator.

---

### 4.2 `Inicializacion.xaml` — Preparación del Entorno

**Tipo:** Sub-Workflow
**Argumentos de salida:** `PathExcel (OutArgument<String>)` *(promovido de variable local en v1.1.0)*

#### Variables Internas

| Variable | Tipo | Propósito |
|---|---|---|
| `fileExists` | `Boolean` | Resultado de la verificación de existencia del archivo Excel |
| `dtHeaders` | `DataTable` | Tabla temporal con encabezados para auto-creación del Excel |

#### Actividades Clave

| Actividad | Configuración |
|---|---|
| `Log Message` | Info: "Iniciando limpieza de entorno y validación de archivos locales" |
| `Kill Process` | Proceso: `EXCEL` — cierra archivos Excel abiertos |
| `Assign` | `PathExcel = Path.Combine(Environment.CurrentDirectory, "Control_Divisas.xlsx")` |
| `File Exists X` | Verifica `PathExcel` → output `fileExists` |
| `If` | Condición: `Not fileExists` |
| `Build DataTable` (rama Then) | Columnas: `Fecha (String)`, `Divisa (String)`, `Valor (String)` |
| `Excel Process Scope X` | Crea el archivo con encabezados si no existe |
| `Log Message` (rama Then) | Info: "Archivo creado automáticamente con encabezados" |

> ✅ **Auto-Bootstrap:** Si `Control_Divisas.xlsx` no existe, se crea automáticamente con encabezados. La ejecución continúa sin interrupción.

---

### 4.3 `Extraer_TRM_USD.xaml` — Extracción Dólar

**Tipo:** Sub-Workflow
**Argumentos de salida:** `strValorUSD (OutArgument<String>)` *(promovido de variable local en v1.1.0)*

> ⚠️ **Estado v1.1.0 — PLACEHOLDER ACTIVO:** El selector UI de `Get Text` sobre Google Finance aún no está configurado. La actividad `Assign` establece `strValorUSD = "5000.00"` para validar el pipeline completo mientras se implementa la UI Automation real.

#### Estructura de Control

```
Sequence
 └── Retry Scope (3 reintentos)
      └── Try-Catch (System.Exception)
           ├── Try:
           │    ├── Log Message [Info]  → "Placeholder for: Use Application/Browser..."
           │    ├── HTTP Request        → ContinueOnError = True  (placeholder)
           │    ├── Assign              → strValorUSD = "5000.00"  ⚠️ PLACEHOLDER
           │    └── Log Message [Info]  → "Valor del Dólar extraído correctamente: 5000.00"
           └── Catch (System.Exception):
                ├── Log Message [Warn]  → "Fallo al intentar extraer USD..."
                └── Rethrow             → activa el siguiente reintento
```

#### Configuración objetivo (cuando se implemente UI Automation real)

| Propiedad | Valor objetivo |
|---|---|
| Actividad | `Use Application/Browser` |
| URL | `https://www.google.com/finance/quote/USD-COP` |
| `InPrivate` | `True` |
| `Close` | `Always` |
| Extracción | `Get Text` sobre el elemento del precio |
| Retry Count | `3` |

---

### 4.4 `Extraer_TRM_EUR.xaml` — Extracción Euro

**Tipo:** Sub-Workflow
**Argumentos de salida:** `strValorEUR (OutArgument<String>)`

> ⚠️ **Estado v1.1.0 — PLACEHOLDER ACTIVO:** Igual que el módulo USD. Asigna `strValorEUR = "4900.00"` de forma fija.

Estructura idéntica al módulo USD con las siguientes diferencias:

| Campo | Valor |
|---|---|
| URL objetivo | `https://www.google.com/finance/quote/EUR-COP` |
| Variable de salida | `strValorEUR` |
| Valor placeholder | `"4900.00"` |
| Fix v1.1.0 aplicado | `Net.HttpStatusCode` → `System.Net.HttpStatusCode` (resuelve error de compilación `BC30456`) |

---

### 4.5 `Limpieza_Validacion_Datos.xaml` — Transformación y Validación

**Tipo:** Sub-Workflow
**Argumentos de entrada:** `strValorUSD (InArgument<String>)`, `strValorEUR (InArgument<String>)`
**Argumentos de salida:** `numValorUSD (OutArgument<Double>)`, `numValorEUR (OutArgument<Double>)`

#### Pipeline de Transformación

**Paso 1 — Limpieza de texto:**
```vb
strValorUSD = If(String.IsNullOrEmpty(strValorUSD), "",
    strValorUSD.Replace("$","").Replace("USD","").Replace(" ","").Replace(",","").Trim())

strValorEUR = If(String.IsNullOrEmpty(strValorEUR), "",
    strValorEUR.Replace("$","").Replace("COP","").Replace(" ","").Replace(",","").Trim())
```
> Elimina símbolos de moneda, espacios y separadores de miles. Soporta formatos americano y colombiano.

**Paso 2 — Conversión numérica:**
```vb
numValorUSD = If(String.IsNullOrEmpty(strValorUSD), 0.0R,
    Double.Parse(strValorUSD, CultureInfo.InvariantCulture))

numValorEUR = If(String.IsNullOrEmpty(strValorEUR), 0.0R,
    Double.Parse(strValorEUR, CultureInfo.InvariantCulture))
```
> `InvariantCulture` garantiza interpretación correcta del punto decimal independientemente de la configuración regional.

**Paso 3 — Validación de Negocio RF-02:**
```vb
' Condición If:
numValorUSD > 0R AndAlso numValorEUR > 0R

' Then:
Log Message [Info] → "Validación RF-02 superada: Datos convertidos correctamente y listos para Excel"

' Else:
Throw New UiPath.Core.BusinessRuleException(
    "Error de integridad: Los valores de las divisas no son numéricos válidos o son menores a cero.")
```

---

### 4.6 `Registro_Excel.xaml` — Persistencia Histórica

**Tipo:** Sub-Workflow
**Argumentos de entrada:** `in_numValorUSD (InArgument<Double>)`, `in_numValorEUR (InArgument<Double>)`, `in_PathExcel (InArgument<String>)`

> **Cambio v1.1.0 — Migración de actividad Excel:**
> La actividad `AppendRange` (Classic) era incompatible con `UiPath.Excel.Activities 3.5.1` causando el error de guardado *"El flujo de trabajo contiene actividades faltantes o no válidas"*. Se reemplazó por `WriteRangeX` (Modern) dentro de `ExcelApplicationCard`.

#### Construcción del DataTable en Memoria

```vb
' Build Data Table → dtDivisas
' Columnas: Fecha (String), Divisa (String), Valor (Double)

' Add Data Row — USD:
new Object() {DateTime.Now.ToString("dd/MM/yyyy"), "USD", in_numValorUSD}

' Add Data Row — EUR:
new Object() {DateTime.Now.ToString("dd/MM/yyyy"), "EUR", in_numValorEUR}
```

#### Configuración de Escritura Excel (RF-04)

| Propiedad | Valor | Razón |
|---|---|---|
| Contenedor | `ExcelApplicationCard` | Requerido por actividades Modern de Excel |
| `WorkbookPath` | `in_PathExcel` | Portabilidad total sin rutas hardcodeadas |
| Actividad de escritura | `WriteRangeX` | Compatible con `UiPath.Excel.Activities 3.5.1+` |
| `Append` | `True` | Agrega filas al final — no sobrescribe historial (RF-04) |
| `ExcludeHeaders` | `True` | Evita duplicar encabezados en cada ejecución |
| `Destination` | `Excel.Sheet("Hoja1")` | Hoja destino fija |
| `Source` | `dtDivisas` | DataTable construido en memoria |

---

## 5. Trazabilidad Completa de Logs

Secuencia de mensajes en una **ejecución exitosa** (v1.1.0 con placeholders activos):

```
[Info]  RPA_Consulta_Divisas_Cash_AA2 ejecución iniciada
[Info]  Iniciando limpieza de entorno y validación de archivos locales

        ↓ Solo si Control_Divisas.xlsx no existe (primera ejecución)
[Warn]  Archivo Control_Divisas.xlsx no encontrado. Creando archivo nuevo...
[Info]  Archivo Control_Divisas.xlsx creado automáticamente con encabezados.

[Info]  Placeholder for: Use Application/Browser (Edge, .../USD-COP) and Get Text.
[Info]  Valor del Dólar extraído correctamente: 5000.00
[Info]  Placeholder for: Use Application/Browser (Edge, .../EUR-COP) and Get Text.
[Info]  Valor del Euro extraído correctamente: 4900.00
[Info]  Validación RF-02 superada: Datos convertidos correctamente y listos para Excel
[Info]  Proceso Cash completado con éxito. Datos guardados en el histórico de Excel.
[Info]  RPA_Consulta_Divisas_Cash_AA2 ejecución terminada in: 00:00:07
```

Mensajes de **alta disponibilidad** (reintento, no fallo definitivo):
```
[Warn]  Fallo al intentar extraer USD. Reintentando...   ← hasta 3 veces
[Warn]  Fallo al intentar extraer EUR. Reintentando...   ← hasta 3 veces
```

Mensajes de **error crítico** (requieren intervención):
```
[Error] BusinessRuleException: Error de integridad: Los valores de las divisas
        no son numéricos válidos o son menores a cero.
```

---

## 6. Requerimientos Funcionales Implementados

| ID | Requerimiento | Flujo que lo implementa | Estado |
|---|---|---|---|
| **RF-01** | Extracción de TRM desde fuente web | `Extraer_TRM_USD.xaml`, `Extraer_TRM_EUR.xaml` | ⚠️ Placeholder activo |
| **RF-02** | Validación de integridad de datos numéricos | `Limpieza_Validacion_Datos.xaml` | ✅ Implementado |
| **RF-03** | Portabilidad de rutas de archivo | `Inicializacion.xaml` (`Path.Combine`) | ✅ Implementado |
| **RF-04** | Registro histórico sin sobrescritura | `Registro_Excel.xaml` (`WriteRangeX` Append=True) | ✅ Implementado |

---

## 7. Dependencias y Paquetes NuGet

| Paquete | Versión | Actividades utilizadas |
|---|---|---|
| `UiPath.UIAutomation.Activities` | Última estable | `Use Application/Browser`, `Get Text`, `Kill Process` |
| `UiPath.Excel.Activities` | **3.5.1** | `ExcelApplicationCard`, `ExcelProcessScopeX`, `WriteRangeX`, `FileExistsX` |
| `UiPath.System.Activities` | Última estable | `Assign`, `If`, `Try-Catch`, `Retry Scope`, `Throw`, `Build DataTable`, `Add Data Row`, `Invoke Workflow File`, `Log Message` |

> ❌ **No usar actividades Classic de Excel** (`AppendRange`, `ReadRange`, `WriteRange` sin sufijo X) con `UiPath.Excel.Activities 3.5.1`. Causan el error *"actividades faltantes o no válidas"* al guardar el flujo de trabajo.

---

## 8. Checklist de Pre-Ejecución

- [ ] **PASO 1** — Cerrar `Control_Divisas.xlsx` si está abierto en Excel.
- [ ] **PASO 2** — Verificar paquetes NuGet: `Proyecto → Administrar Paquetes` sin advertencias ⚠️.
- [ ] **PASO 3** — Confirmar acceso a `https://www.google.com/finance/quote/USD-COP` y `EUR-COP`.
- [ ] **PASO 4 (Pendiente — RF-01)** — Usar **"Indicate in App"** en `Extraer_TRM_USD.xaml` y `Extraer_TRM_EUR.xaml` para apuntar al elemento del precio en Google Finance y reemplazar los placeholders.
- [ ] **PASO 5** — Ejecutar `Diseño → Analizar proyecto` — cero errores antes de correr.

---

## 9. Mantenimiento y Consideraciones de Producción

### Escenarios de Fallo y Acciones

| Escenario | Comportamiento del Robot | Acción Requerida |
|---|---|---|
| Google Finance cambia estructura HTML | `Get Text` falla → 3 reintentos → excepción | Re-indicar el selector con "Indicate in App" |
| Archivo Excel abierto por otro usuario | `WriteRangeX` falla dentro de `ExcelApplicationCard` | Cerrar el archivo y re-ejecutar |
| Sin conexión a internet | `Use Application/Browser` falla → 3 reintentos | Verificar red y re-ejecutar |
| Valor web en formato inesperado | `Double.Parse` falla → `BusinessRuleException` (RF-02) | Revisar texto capturado y ajustar `.Replace()` |
| Primera ejecución en máquina nueva | Excel no existe → Auto-creado por el robot | Ninguna — proceso autónomo |
| Paquete Excel en versión Classic | `AppendRange` / `WriteRange` no encontrados | Actualizar a `UiPath.Excel.Activities 3.5.1+` y usar `WriteRangeX` |

### Buenas Prácticas para Operación

1. **Programación en Orchestrator:** Ejecución automática diaria (ej. 8:00 AM, hora Colombia).
2. **Monitoreo:** Filtrar logs por `[Error]` y `[Warn]` en Orchestrator tras cada ejecución.
3. **Backup del Excel:** Copia semanal automática de `Control_Divisas.xlsx`.
4. **Selector de UI:** Si Google actualiza su diseño, solo re-indicar `Get Text` en los dos módulos de extracción; el resto del pipeline no requiere cambios.
5. **Escalabilidad:** Para agregar nuevas divisas (GBP, JPY, etc.), duplicar `Extraer_TRM_USD.xaml`, cambiar URL y variable de salida, y agregar el `Invoke` en `Main.xaml` con su `OutArgument` correspondiente.

---

## 10. Historial de Cambios

### v1.1.0 — 09/05/2026 *(versión actual)*

**Correcciones críticas — pipeline completamente funcional:**

| Archivo modificado | Cambio aplicado |
|---|---|
| `Registro_Excel.xaml` | ❌ `AppendRange` (Classic, incompatible con Excel Activities 3.5.1) → ✅ `WriteRangeX` dentro de `ExcelApplicationCard` (Modern). Resuelve el error de guardado *"actividades faltantes o no válidas"*. |
| `Extraer_TRM_USD.xaml` | `strValorUSD` promovida de **variable local** a **`OutArgument<String>`** para ser accesible desde `Main.xaml`. Agrega `Assign strValorUSD = "5000.00"` (placeholder). |
| `Extraer_TRM_EUR.xaml` | `strValorEUR` ya era `OutArgument`; se agrega `Assign strValorEUR = "4900.00"` (placeholder). Corregida referencia `Net.HttpStatusCode` → `System.Net.HttpStatusCode` (resuelve `BC30456`). |
| `Inicializacion.xaml` | `PathExcel` promovida a **`OutArgument<String>`**. Agrega `Assign` con ruta al Excel de control. |
| `Main.xaml` | Los 5 `InvokeWorkflowFile` tenían diccionarios de argumentos **vacíos**. Se poblan con el cableado completo de `InArgument` y `OutArgument` para los 10 parámetros del pipeline. |

**Resultado validado:** Ejecución completa del pipeline en **7 segundos** sin errores de validación.

---

### v1.0.0 — 05/05/2026

- Versión inicial del proyecto.
- Creación de los 6 módulos `.xaml` con estructura base.
- Definición de argumentos en `Limpieza_Validacion_Datos.xaml` y `Registro_Excel.xaml`.
- Auto-bootstrap de `Control_Divisas.xlsx` en `Inicializacion.xaml`.
- `Main.xaml` con orquestación básica (argumentos sin cablear — corregido en v1.1.0).

---

*Proyecto Cash Divisas AA2 — Ignacio Javier Julio Posada*
*Última actualización: 09/05/2026 | Versión del documento: 1.1.0*
