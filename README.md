# Automatización de Reportes de Licencias Médicas - SLEP Los Libertadores

Este proyecto surge con el objetivo de modernizar y automatizar la gestión de licencias médicas de la organización SLEP Los Libertadores.

## El Problema

La organización dependía de un sistema de registro basado en hojas de cálculo (Excel) que presentaba problemas de escalabilidad y calidad de datos:

- Incoherencias: Datos redundantes y contradictorios.
- Errores de entrada: Múltiples errores ortográficos y falta de estandarización en los nombres.
- Procesos manuales: El reporte de licencias y la imputación de datos requerían un esfuerzo manual exhaustivo y propenso a errores.

## La Solución

Se desarrolló una plataforma web que actúa como un procesador de datos inteligente. La herramienta carga la base de datos histórica (desde 2024) y transforma el flujo de trabajo mediante:

1. Normalización de Datos: Limpieza automática de errores ortográficos y estandarización de entradas.
2. Validaciones Cruzadas: El sistema verifica la integridad de la información entre distintas tablas (establecimientos, AFPs, etc.) para detectar anomalías.
3. Generación de Reportes Automatizados: En lugar de editar el Excel manualmente, la herramienta genera un archivo ZIP que incluye:
    - Reportes con Alertas: Identificación inmediata de inconsistencias que requieren atención.
    - Planillas Normalizadas: Documentos listos para su uso, con campos autorrellenados para facilitar la imputación final.

## Stack Tecnológico

- Frontend & Documentación: [Quarto](https://quarto.org/) (HTML, CSS, JavaScript) para la interfaz de usuario y la documentación técnica.
- Procesamiento de Datos: Python (ubicado en `/scripts`) para la lógica de migración y limpieza.
- Análisis de Datos: Power BI (`.pbit`) para la visualización de indicadores clave (KPIs) de las licencias.
- Almacenamiento: Excel (`.xlsx`) como fuente de datos y formato de salida.

## Estructura del Repositorio

- `index.qmd`: Página de inicio y documentación del proyecto.
- `migrador.qmd`: Documentación y herramienta de migración de datos.
- `assets/`: Recursos estáticos, incluyendo el dashboard de Power BI y la lógica de JS (`ui.js`, `worker.js`).
- `scripts/slep/`: Núcleo de procesamiento en Python (`core.py`, `utils.py`).
- `assets/tables/`: Tablas maestras de referencia (Establecimientos, AFPs).
