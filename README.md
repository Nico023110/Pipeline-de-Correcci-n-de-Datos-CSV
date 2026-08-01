# Pipeline de Corrección de Datos CSV (RIPS / Facturación)

Motor automatizado para la validación, limpieza, estandarización y corrección de inconsistencias en archivos CSV de RIPS y facturación de servicios de salud en Colombia.

## 📁 Estructura del Proyecto

```
correccion_csv/
├── data/                         # Tablas de equivalencias (CUPS, Finalidades, CUM/IUM)
├── factura_corregida/            # Archivos CSV corregidos de salida
├── fuente/                       # Archivos CSV fuente originales por EAPB/EPS
├── historial/                    # Histórico de archivos procesados y respaldos
├── log_errores/                  # Reportes y logs consolidados de inconsistencias en Excel
└── scripts/
    ├── motor_correccion.py       # Motor principal de validación y corrección unitario
    ├── motor_correccion_multiple.py # Motor de procesamiento masivo por lotes
    ├── generar_reporte_consolidado_eapb.py # Generación de reportes consolidados por EAPB
    ├── interfaz_rips.py          # Interfaz gráfica/consola para selección de archivos RIPS
    └── PRUEBA V1/V2/V3.py        # Scripts de pruebas e iteraciones de desarrollo
```

## 🚀 Requisitos e Instalación

### Requisitos Previos
* Python 3.8+

### Librerías Requeridas
```bash
pip install pandas numpy openpyxl
```

## ⚙️ Uso

### Ejecutar Corrección de RIPS
```bash
python scripts/motor_correccion.py
```

### Ejecutar Procesamiento Masivo
```bash
python scripts/motor_correccion_multiple.py
```

### Generar Reporte Consolidado por EAPB
```bash
python scripts/generar_reporte_consolidado_eapb.py
```

## 🔍 Funcionalidades Principales

* Validación de campos obligatorios según la norma RIPS.
* Cruce y corrección automática de códigos CUPS contra finalidades.
* Normalización de identificadores CUM/IUM en medicamentos.
* Corrección de formato de fechas, causas externas y tipos de documento.
* Generación de logs detallados de inconsistencias en Excel por cada factura.
