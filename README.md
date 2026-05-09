# 🤖 RPA Consulta Divisas — Cash AA2

> Automatización robótica desarrollada en **UiPath Studio (Modern Design)** para la empresa de divisas **Cash**.
> Extrae, valida y registra de forma autónoma la TRM del Dólar (USD) y el Euro (EUR) desde Google Finance.

---

## 📋 Descripción

Este proceso RPA realiza diariamente las siguientes operaciones de forma completamente autónoma:

1. **Limpieza del entorno** — Cierra instancias abiertas de Excel para evitar conflictos.
2. **Verificación y auto-creación** del archivo maestro `Control_Divisas.xlsx` si no existe.
3. **Extracción web** — Navega a Google Finance en modo InPrivate (sin cookies) y captura el precio actual de USD/COP y EUR/COP.
4. **Transformación y validación** — Limpia los datos crudos, los convierte a tipo numérico con cultura invariante y valida su integridad (RF-02).
5. **Registro histórico** — Escribe los datos validados al final del archivo Excel sin sobrescribir el historial previo.

> ⚠️ **Estado actual:** Los módulos de extracción web (`Extraer_TRM_USD.xaml` y `Extraer_TRM_EUR.xaml`) funcionan con **valores placeholder** (`5000.00` USD / `4900.00` EUR) mientras se configura la automatización UI real sobre Google Finance. El resto del pipeline (validación + registro Excel) está **completamente funcional**.

---

## 🗂️ Estructura del Proyecto

```
RPA_Consulta_Divisas_Cash_AA2/
│
├── Main.xaml                        ← Orquestador maestro (entry point)
├── Inicializacion.xaml              ← Limpieza del entorno y validación de archivos
├── Extraer_TRM_USD.xaml             ← Extracción TRM Dólar (USD/COP) [placeholder]
├── Extraer_TRM_EUR.xaml             ← Extracción TRM Euro (EUR/COP)  [placeholder]
├── Limpieza_Validacion_Datos.xaml   ← Transformación y validación RF-02
├── Registro_Excel.xaml              ← Escritura histórica en Excel (RF-04)
│
├── Control_Divisas.xlsx             ← Archivo de datos históricos (auto-creado)
├── project.json                     ← Configuración del proyecto UiPath
├── project.uiproj                   ← Archivo de solución UiPath
├── DOCUMENTACION_TECNICA.md         ← Documentación técnica completa
└── README.md                        ← Este archivo
```

---

## ⚙️ Requisitos Previos

| Requisito | Detalle |
|---|---|
| UiPath Studio | Modern Design — Windows Framework |
| `UiPath.UIAutomation.Activities` | Última versión estable |
| `UiPath.Excel.Activities` | `3.5.1` o superior |
| `UiPath.System.Activities` | Última versión estable |
| Microsoft Edge | Instalado (para navegación InPrivate) |
| Conectividad | Acceso a `https://www.google.com/finance` |

---

## 🚀 Cómo ejecutar

### Primera ejecución
1. Clona el repositorio:
   ```bash
   git clone git@github.com:ignaciojulio/RPA_Consulta_Divisas_Cash_AA2.git
   ```
2. Abre la carpeta del proyecto en **UiPath Studio**.
3. Ve a `Proyecto → Administrar Paquetes` y verifica que los 3 paquetes estén instalados.
4. *(Pendiente)* Abre `Extraer_TRM_USD.xaml` y usa **"Indicate in App"** en la actividad `Get Text` para apuntar al precio del Dólar en Google Finance.
5. *(Pendiente)* Repite el paso anterior en `Extraer_TRM_EUR.xaml` para el Euro.
6. Abre `Main.xaml` y presiona ▶ **Run**.

> ✅ El archivo `Control_Divisas.xlsx` se creará automáticamente en la primera ejecución.

---

## 🏗️ Arquitectura y Flujo de Datos

```
Main.xaml  (Try-Catch Global)
 │
 ├── Inicializacion.xaml              OUT → PathExcel (String)
 ├── Extraer_TRM_USD.xaml             OUT → strValorUSD (String)
 ├── Extraer_TRM_EUR.xaml             OUT → strValorEUR (String)
 ├── Limpieza_Validacion_Datos.xaml   IN  → strValorUSD, strValorEUR
 │                                    OUT → numValorUSD (Double), numValorEUR (Double)
 └── Registro_Excel.xaml              IN  → numValorUSD, numValorEUR, PathExcel
```

> Todos los `InvokeWorkflowFile` tienen sus argumentos correctamente cableados desde el 09/05/2026.

---

## 📊 Requerimientos Funcionales

| ID | Descripción | Estado |
|---|---|---|
| RF-01 | Extracción de TRM desde Google Finance | ⚠️ Placeholder activo |
| RF-02 | Validación de integridad de datos numéricos | ✅ Implementado |
| RF-03 | Portabilidad de rutas relativas | ✅ Implementado |
| RF-04 | Registro histórico sin sobrescritura | ✅ Implementado (`WriteRangeX`) |

---

## 🛡️ Características de Alta Disponibilidad

- **Retry Scope** — 3 reintentos automáticos en cada extracción web.
- **Try-Catch** — Manejo de errores en extracción web y escritura Excel.
- **InPrivate Mode** — Sesiones Edge limpias, sin cookies ni historial.
- **Auto-Bootstrap** — El robot crea su propio Excel si no lo encuentra.
- **Rethrow al Orchestrator** — Propagación elegante de errores críticos.

---

## 📁 Estructura del Archivo de Datos

El robot genera y mantiene `Control_Divisas.xlsx` con la siguiente estructura:

| Fecha | Divisa | Valor |
|---|---|---|
| 09/05/2026 | USD | 5000.00 |
| 09/05/2026 | EUR | 4900.00 |

> ⚠️ El archivo `.xlsx` está excluido del repositorio (`.gitignore`) para proteger los datos históricos.

---

## 📝 Historial de Cambios

| Versión | Fecha | Descripción |
|---|---|---|
| **1.1.0** | 09/05/2026 | Fix cableado de argumentos en `Main.xaml`; `strValorUSD`/`strValorEUR`/`PathExcel` promovidos a `OutArgument`; `Registro_Excel` migrado a `WriteRangeX` + `ExcelApplicationCard`; fix `System.Net.HttpStatusCode` en EUR |
| 1.0.0 | 05/05/2026 | Versión inicial del proyecto |

---

## 👨‍💻 Autor

**Ignacio Javier Julio Posada**
📧 ignacioj.julio@gmail.com
🏫 Proyecto desarrollado para certificación SENA — Automatización Robótica de Procesos (RPA)

---

## 📄 Licencia

Proyecto de uso educativo y corporativo interno para la empresa **Cash**.
Todos los derechos reservados © 2026.
