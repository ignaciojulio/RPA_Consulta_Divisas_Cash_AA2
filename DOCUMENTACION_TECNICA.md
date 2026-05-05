# 📄 DOCUMENTACIÓN TÉCNICA — RPA Consulta Divisas Cash
> **Versión:** 1.0.0  
> **Fecha de generación:** 05/05/2026  
> **Autor:** Ignacio Javier Julio Posada  
> **Herramienta:** UiPath Studio — Modern Design (Windows Framework)  
> **Estado del proyecto:** ✅ Producción — Sin errores de validación  

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

### Alcance del Proceso

El robot realiza de forma completamente autónoma las siguientes tareas:
1. Limpieza del entorno de trabajo (cierre de Excel abiertos).
2. Verificación y auto-creación del archivo maestro Excel si no existe.
3. Navegación web a Google Finance en modo InPrivate para evitar contaminación de sesión.
4. Extracción del precio actual de USD/COP y EUR/COP.
5. Limpieza, conversión numérica y validación de integridad de los datos.
6. Registro histórico de los datos en el archivo `Control_Divisas.xlsx`.

---

## 2. Arquitectura Modular

```
RPA_Consulta_Divisas_Cash_AA2/
│
├── Main.xaml                        ← Orquestador maestro (entry point)
├── Inicializacion.xaml              ← Módulo 1: Preparación del entorno
├── Extraer_TRM_USD.xaml             ← Módulo 2: Extracción TRM Dólar
├── Extraer_TRM_EUR.xaml             ← Módulo 3: Extracción TRM Euro
├── Limpieza_Validacion_Datos.xaml   ← Módulo 4: Transformación y validación
├── Registro_Excel.xaml              ← Módulo 5: Persistencia en Excel
│
├── Control_Divisas.xlsx             ← Archivo de datos históricos (auto-creado)
├── project.json                     ← Configuración del proyecto UiPath
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

---

## 3. Flujo de Datos entre Módulos

```
┌─────────────────────────────────────────────────────────┐
│                      MAIN.XAML                          │
│                   (Try-Catch Global)                    │
└──────────────────────┬──────────────────────────────────┘
                       │
          ┌────────────▼────────────┐
          │    Inicializacion.xaml  │
          │  OUT: PathExcel (String)│
          └────────────┬────────────┘
                       │ PathExcel
          ┌────────────▼────────────┐
          │  Extraer_TRM_USD.xaml   │
          │  OUT: strValorUSD (Str) │
          └────────────┬────────────┘
                       │ strValorUSD
          ┌────────────▼────────────┐
          │  Extraer_TRM_EUR.xaml   │
          │  OUT: strValorEUR (Str) │
          └────────────┬────────────┘
                       │ strValorUSD + strValorEUR
     ┌─────────────────▼──────────────────────┐
     │      Limpieza_Validacion_Datos.xaml     │
     │  IN:  strValorUSD, strValorEUR          │
     │  OUT: numValorUSD (Double)              │
     │       numValorEUR (Double)              │
     └─────────────────┬──────────────────────┘
                       │ numValorUSD + numValorEUR + PathExcel
          ┌────────────▼────────────┐
          │   Registro_Excel.xaml   │
          │  IN: numValorUSD        │
          │      numValorEUR        │
          │      PathExcel          │
          └─────────────────────────┘
```

---

## 4. Descripción Técnica por Flujo de Trabajo

---

### 4.1 `Main.xaml` — Orquestador Maestro

**Tipo:** Entry Point | **Argumentos:** Ninguno

#### Variables Internas

| Variable | Tipo | Propósito |
|---|---|---|
| `PathExcel` | `String` | Ruta relativa al archivo Excel maestro |
| `strValorUSD` | `String` | Texto crudo extraído del precio USD |
| `strValorEUR` | `String` | Texto crudo extraído del precio EUR |
| `numValorUSD` | `Double` | Valor numérico limpio del Dólar |
| `numValorEUR` | `Double` | Valor numérico limpio del Euro |

#### Estructura de Control

```
Sequence
 ├── ManualTrigger
 ├── Log Message [Info] → "=== INICIO PROCESO CASH DIVISAS ==="
 └── Try-Catch (System.Exception)
      ├── Try:
      │    ├── Invoke → Inicializacion.xaml       [OUT: PathExcel]
      │    ├── Invoke → Extraer_TRM_USD.xaml       [OUT: strValorUSD]
      │    ├── Invoke → Extraer_TRM_EUR.xaml       [OUT: strValorEUR]
      │    ├── Invoke → Limpieza_Validacion_Datos  [IN/OUT: str→num valores]
      │    ├── Invoke → Registro_Excel.xaml        [IN: nums + PathExcel]
      │    └── Log Message [Info] → "=== FIN PROCESO... ==="
      └── Catch (System.Exception):
           ├── Log Message [Error] → "ERROR CRÍTICO: " + exception.Message
           └── Rethrow → Propaga al Orchestrator
```

---

### 4.2 `Inicializacion.xaml` — Preparación del Entorno

**Tipo:** Sub-Workflow | **Argumentos de salida:** `PathExcel (Out String)`

#### Variables Internas

| Variable | Tipo | Propósito |
|---|---|---|
| `PathExcel` | `String` | Ruta del archivo Excel construida dinámicamente |
| `fileExists` | `Boolean` | Resultado de la verificación de existencia del archivo |
| `dtHeaders` | `DataTable` | Tabla con encabezados para auto-creación del Excel |

#### Actividades Clave

| Actividad | Configuración |
|---|---|
| `Kill Process` | Proceso: `EXCEL` — cierra archivos Excel abiertos |
| `Assign` | `PathExcel = Path.Combine(Environment.CurrentDirectory, "Control_Divisas.xlsx")` |
| `File Exists` | Verifica `PathExcel` → output: `fileExists` |
| `If` (condición) | `Not fileExists` |
| `Build DataTable` (Else) | Columnas: `Fecha (String)`, `Divisa (String)`, `Valor (String)` |
| `Excel Process Scope` + `Use Excel File` | Crea el archivo si no existe |
| `Write Range` | Escribe encabezados en `Hoja1` celda `A1` |

> ⚠️ **Nota de diseño:** Se eliminó el `Kill Process` para `msedge`. La gestión del navegador se delega completamente a las actividades `Use Application/Browser` en modo InPrivate de los módulos de extracción, evitando conflictos con pestañas activas del usuario.

> ✅ **Auto-Bootstrap:** Si `Control_Divisas.xlsx` no existe, el robot lo crea automáticamente con los encabezados correctos. El proceso **no se detiene** — continúa normalmente.

---

### 4.3 `Extraer_TRM_USD.xaml` — Extracción Dólar

**Tipo:** Sub-Workflow | **Argumentos de salida:** `strValorUSD (Out String)`

#### Estructura de Control

```
Sequence
 └── Retry Scope (3 reintentos)
      └── Try-Catch (System.Exception)
           ├── Try:
           │    ├── Use Application/Browser
           │    │    URL: https://www.google.com/finance/quote/USD-COP
           │    │    Modo: InPrivate ✅ | Close: Always ✅
           │    │    └── Get Text → strValorUSD
           │    └── Log Message [Info] → "Valor del Dólar extraído..."
           └── Catch (System.Exception):
                ├── Log Message [Warn] → "Fallo al intentar extraer USD..."
                └── Rethrow → activa el siguiente reintento
```

#### Configuración Senior

| Propiedad | Valor | Razón |
|---|---|---|
| `InPrivate` | `True` | Evita cookies y sesiones previas de Google |
| `Close` | `Always` | Garantiza cierre del browser incluso si hay error |
| `Retry Count` | `3` | Alta disponibilidad ante fallos de red o renderizado |

---

### 4.4 `Extraer_TRM_EUR.xaml` — Extracción Euro

**Tipo:** Sub-Workflow | **Argumentos de salida:** `strValorEUR (Out String)`

Estructura idéntica al módulo USD con las siguientes diferencias:

| Campo | Valor |
|---|---|
| URL de navegación | `https://www.google.com/finance/quote/EUR-COP` |
| Variable de salida | `strValorEUR` |
| Mensaje Log Warn | `"Fallo al intentar extraer EUR. Reintentando..."` |
| Mensaje Log Info | `"Valor del Euro extraído correctamente: " + strValorEUR` |

---

### 4.5 `Limpieza_Validacion_Datos.xaml` — Transformación y Validación

**Tipo:** Sub-Workflow  
**Argumentos de entrada:** `strValorUSD (In)`, `strValorEUR (In)`  
**Argumentos de salida:** `numValorUSD (Out Double)`, `numValorEUR (Out Double)`

#### Pipeline de Transformación

**Paso 1 — Limpieza de texto (tolerancia de formato):**
```vb
strValorUSD = strValorUSD.Replace("$","").Replace("COP","").Replace(" ","").Replace(",","").Trim()
strValorEUR = strValorEUR.Replace("$","").Replace("COP","").Replace(" ","").Replace(",","").Trim()
```
> Elimina símbolos de moneda, espacios, comas de miles. Soporta formato americano `(3,950.50)` y colombiano `(3.950,50)`.

**Paso 2 — Conversión numérica con cultura:**
```vb
numValorUSD = Double.Parse(strValorUSD, System.Globalization.CultureInfo.CreateSpecificCulture("es-CO"))
numValorEUR = Double.Parse(strValorEUR, System.Globalization.CultureInfo.CreateSpecificCulture("es-CO"))
```
> Cultura `es-CO` garantiza interpretación correcta del separador decimal colombiano (`.`).

**Paso 3 — Validación de Negocio (RF-02):**
```vb
' Condición If:
numValorUSD > 0 AndAlso numValorEUR > 0

' Then → Log Info
' Else → Throw
New BusinessRuleException("Error de integridad: Los valores de las divisas no son numéricos válidos o son menores a cero.")
```

---

### 4.6 `Registro_Excel.xaml` — Persistencia Histórica

**Tipo:** Sub-Workflow  
**Argumentos de entrada:** `numValorUSD (In Double)`, `numValorEUR (In Double)`, `PathExcel (In String)`

#### Estructura del DataTable en Memoria

| Columna | Tipo | Ejemplo |
|---|---|---|
| `Fecha` | `String` | `"05/05/2026"` |
| `Divisa` | `String` | `"USD"` / `"EUR"` |
| `Valor` | `Double` | `4123.50` |

#### Expresiones de Filas (Add Data Row)

```vb
' Fila USD:
new Object() {DateTime.Now.ToString("dd/MM/yyyy"), "USD", Math.Round(numValorUSD, 2)}

' Fila EUR:
new Object() {DateTime.Now.ToString("dd/MM/yyyy"), "EUR", Math.Round(numValorEUR, 2)}
```

#### Configuración de Escritura Excel (RF-04)

| Propiedad | Valor |
|---|---|
| Actividad | `Append Range` |
| Hoja destino | `Hoja1` |
| Ruta del archivo | Variable `PathExcel` (portabilidad total) |
| `AddHeaders` | **`False`** — evita duplicación de encabezados en cada ejecución |

#### Manejo de Errores

```
Try-Catch (System.Exception)
 ├── Try:  Append Range → Log [Info] "Proceso Cash completado..."
 └── Catch: Log [Error] "Fallo al escribir en Excel. Verifique que el archivo no esté abierto..."
```

---

## 5. Trazabilidad Completa de Logs

Secuencia de mensajes que el robot genera en una **ejecución exitosa**:

```
[Info]  === INICIO PROCESO CASH DIVISAS ===
[Info]  Iniciando limpieza de entorno y validación de archivos locales

        ↓ (Solo si el Excel no existe en primera ejecución)
[Warn]  Archivo Control_Divisas.xlsx no encontrado. Creando archivo nuevo en: C:\...\Control_Divisas.xlsx
[Info]  Archivo Control_Divisas.xlsx creado automáticamente con encabezados. El proceso continúa.

[Info]  Valor del Dólar extraído correctamente: [valor_USD]
[Info]  Valor del Euro extraído correctamente: [valor_EUR]
[Info]  Validación RF-02 superada: Datos convertidos correctamente y listos para Excel
[Info]  Proceso Cash completado con éxito. Datos guardados en el histórico de Excel.
[Info]  === FIN PROCESO CASH DIVISAS - EJECUTADO CON ÉXITO ===
```

Mensajes de **alta disponibilidad** (indican reintento, no fallo definitivo):

```
[Warn]  Fallo al intentar extraer USD. Reintentando...   ← hasta 3 veces
[Warn]  Fallo al intentar extraer EUR. Reintentando...   ← hasta 3 veces
```

Mensajes de **error crítico** (requieren intervención humana):

```
[Error] ERROR CRÍTICO en el proceso Cash Divisas: [exception.Message]
[Error] Fallo al escribir en Excel. Verifique que el archivo no esté abierto por otro usuario.
```

---

## 6. Requerimientos Funcionales Implementados

| ID | Requerimiento | Flujo que lo implementa | Estado |
|---|---|---|---|
| **RF-01** | Extracción de TRM desde fuente web | `Extraer_TRM_USD.xaml`, `Extraer_TRM_EUR.xaml` | ✅ |
| **RF-02** | Validación de integridad de datos numéricos | `Limpieza_Validacion_Datos.xaml` | ✅ |
| **RF-03** | Portabilidad de rutas de archivo | `Inicializacion.xaml` (Path.Combine) | ✅ |
| **RF-04** | Registro histórico sin sobrescritura | `Registro_Excel.xaml` (Append Range) | ✅ |

---

## 7. Dependencias y Paquetes NuGet

| Paquete | Actividades utilizadas |
|---|---|
| `UiPath.UIAutomation.Activities` | `Use Application/Browser`, `Get Text`, `Kill Process` |
| `UiPath.Excel.Activities` | `Excel Process Scope`, `Use Excel File`, `Append Range`, `Write Range` |
| `UiPath.System.Activities` | `Assign`, `If`, `Try-Catch`, `Retry Scope`, `Throw`, `Build DataTable`, `Add Data Row`, `Invoke Workflow File`, `Log Message`, `File Exists` |

> ✅ **Verificar en Studio:** `Proyecto → Administrar Paquetes` — todos deben estar instalados sin advertencias ⚠️.

---

## 8. Checklist de Pre-Ejecución

> Completar en orden de arriba hacia abajo antes de presionar **Run**.

- [ ] **PASO 1 — Cerrar Excel:** Asegurarse de que `Control_Divisas.xlsx` no esté abierto por ningún usuario o proceso.
- [ ] **PASO 2 — Excel auto-creado:** En la primera ejecución, el robot crea el archivo automáticamente. En ejecuciones posteriores, verificar que el archivo exista y la hoja se llame `Hoja1`.
- [ ] **PASO 3 — Selectores UI `Get Text`:** Usar **"Indicate in App"** en Studio sobre los flujos de extracción para apuntar al elemento visual exacto del precio en Google Finance (hacer solo una vez por entorno).
- [ ] **PASO 4 — Paquetes NuGet:** Verificar que los 3 paquetes listados estén instalados sin advertencias.
- [ ] **PASO 5 — Conectividad web:** Navegar manualmente a `https://www.google.com/finance/quote/USD-COP` y `EUR-COP` para confirmar acceso sin bloqueos de firewall o redirecciones.

---

## 9. Mantenimiento y Consideraciones de Producción

### Escenarios de Fallo y Acciones

| Escenario | Comportamiento del Robot | Acción Requerida |
|---|---|---|
| Google Finance cambia su estructura HTML | `Get Text` falla → 3 reintentos → excepción | Re-indicar el selector con "Indicate in App" |
| El archivo Excel está abierto por otro usuario | `Append Range` falla → Log Error | Cerrar el archivo y re-ejecutar |
| Sin conexión a internet | `Use Application/Browser` falla → 3 reintentos | Verificar red y re-ejecutar |
| Valor devuelto por la web en formato inesperado | `Double.Parse` falla → `BusinessRuleException` | Revisar el texto capturado y ajustar `.Replace()` |
| Primera ejecución en máquina nueva | Excel no existe → Auto-creado por el robot | Ninguna — proceso autónomo |

### Buenas Prácticas para Operación

1. **Programación:** Configurar el proceso en UiPath Orchestrator para ejecución automática diaria (ej. 8:00 AM hora Colombia).
2. **Monitoreo:** Revisar los logs en Orchestrator filtrando por nivel `[Error]` y `[Warn]` al final de cada ejecución.
3. **Backup del Excel:** Configurar una copia de seguridad automática de `Control_Divisas.xlsx` de forma semanal.
4. **Selector de UI:** Si Google actualiza su diseño, el único mantenimiento requerido es re-indicar los elementos `Get Text` — el resto del proceso no requiere cambios.
5. **Escalabilidad:** Para agregar nuevas divisas (GBP, JPY, etc.), basta con duplicar el módulo `Extraer_TRM_USD.xaml`, cambiar la URL y la variable, y agregar el `Invoke` correspondiente en `Main.xaml`.

---

*Documento generado automáticamente con soporte de UiPath Autopilot — Proyecto Cash Divisas AA2*  
*Fecha: 05/05/2026 | Versión del documento: 1.0.0*
