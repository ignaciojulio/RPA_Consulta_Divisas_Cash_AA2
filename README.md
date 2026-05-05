# 🤖 RPA Consulta Divisas — Cash AA2

> Automatización robótica desarrollada en **UiPath Studio (Modern Design)** para la empresa de divisas **Cash**.
> Extrae, valida y registra de forma autónoma la TRM del Dólar (USD) y el Euro (EUR) desde Google Finance.

---

## 📋 Descripción

Este proceso RPA realiza diariamente las siguientes operaciones de forma completamente autónoma:

1. **Limpieza del entorno** — Cierra instancias abiertas de Excel para evitar conflictos.
2. **Verificación y auto-creación** del archivo maestro `Control_Divisas.xlsx` si no existe.
3. **Extracción web** — Navega a Google Finance en modo InPrivate (sin cookies) y captura el precio actual de USD/COP y EUR/COP.
4. **Transformación y validación** — Limpia los datos crudos, los convierte a tipo numérico con cultura `es-CO` y valida su integridad (RF-02).
5. **Registro histórico** — Escribe los datos validados al final del archivo Excel sin sobrescribir el historial previo.

---

## 🗂️ Estructura del Proyecto

```
RPA_Consulta_Divisas_Cash_AA2/
│
├── Main.xaml                        ← Orquestador maestro (entry point)
├── Inicializacion.xaml              ← Limpieza del entorno y validación de archivos
├── Extraer_TRM_USD.xaml             ← Extracción TRM Dólar (USD/COP)
├── Extraer_TRM_EUR.xaml             ← Extracción TRM Euro (EUR/COP)
├── Limpieza_Validacion_Datos.xaml   ← Transformación y validación RF-02
├── Registro_Excel.xaml              ← Escritura histórica en Excel (RF-04)
│
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
| `UiPath.Excel.Activities` | Última versión estable |
| `UiPath.System.Activities` | Última versión estable |
| Microsoft Edge | Instalado (para navegación InPrivate) |
| Conectividad | Acceso a `https://www.google.com/finance` |

---

## 🚀 Cómo ejecutar

### Primera ejecución
1. Clona el repositorio:
   ```bash
   git clone https://github.com/TU_USUARIO/RPA_Consulta_Divisas_Cash_AA2.git
   ```
2. Abre la carpeta del proyecto en **UiPath Studio**.
3. Ve a `Proyecto → Administrar Paquetes` y verifica que los 3 paquetes estén instalados.
4. Abre `Extraer_TRM_USD.xaml` y usa **"Indicate in App"** en la actividad `Get Text` para apuntar al precio del Dólar en Google Finance.
5. Repite el paso anterior en `Extraer_TRM_EUR.xaml` para el Euro.
6. Abre `Main.xaml` y presiona ▶ **Run**.

> ✅ El archivo `Control_Divisas.xlsx` se creará automáticamente en la primera ejecución.

---

## 🏗️ Arquitectura y Flujo de Datos

```
Main.xaml  (Try-Catch Global)
 │
 ├── Inicializacion.xaml              OUT → PathExcel
 ├── Extraer_TRM_USD.xaml             OUT → strValorUSD
 ├── Extraer_TRM_EUR.xaml             OUT → strValorEUR
 ├── Limpieza_Validacion_Datos.xaml   IN  → strValorUSD, strValorEUR
 │                                    OUT → numValorUSD, numValorEUR
 └── Registro_Excel.xaml              IN  → numValorUSD, numValorEUR, PathExcel
```

---

## 📊 Requerimientos Funcionales

| ID | Descripción | Estado |
|---|---|---|
| RF-01 | Extracción de TRM desde Google Finance | ✅ |
| RF-02 | Validación de integridad de datos numéricos | ✅ |
| RF-03 | Portabilidad de rutas relativas | ✅ |
| RF-04 | Registro histórico sin sobrescritura | ✅ |

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
| 05/05/2026 | USD | 4123.50 |
| 05/05/2026 | EUR | 4456.75 |

> ⚠️ El archivo `.xlsx` está excluido del repositorio (`.gitignore`) para proteger los datos históricos.

---

## 👨‍💻 Autor

**Ignacio Javier Julio Posada**
📧 ignacioj.julio@gmail.com
🏫 Proyecto desarrollado para certificación SENA — Automatización Robótica de Procesos (RPA)

---

## 📄 Licencia

Proyecto de uso educativo y corporativo interno para la empresa **Cash**.
Todos los derechos reservados © 2026.
