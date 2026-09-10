# Dashboard Analítico: Mercado Laboral Tecnológico (2020-2026)

Este proyecto consiste en un panel interactivo (Dashboard) diseñado para analizar métricas cuantitativas y descriptivas del mercado digital global. Presenta las tendencias de compensación salarial, modalidades de trabajo y demanda de competencias técnicas en el sector tecnológico.

## 🚀 Tecnologías Utilizadas
* **HTML5 & CSS3**: Arquitectura responsiva y diseño en modo oscuro (UI/UX) optimizado para legibilidad analítica.
* **D3.js (v7)**: Manipulación de DOM impulsada por datos para la renderización de gráficos vectoriales (SVG) de alta fidelidad.
* **Google Charts**: Renderizado de mapas geoespaciales interactivos.

## 📊 Origen y Procesamiento de Datos (ETL)
Los datos han sido pre-procesados mediante Python (Pandas) y consolidados a partir de las siguientes fuentes sintéticas y reales:
1. `kaggle_cleaned.csv`: Evolución salarial histórica (2020-2026).
2. `encuesta_sysarmy_20261_CSV.csv`: Ajustes regionales y modalidades.
3. `encuesta_stackoverflow_2025_CSV.csv`: Estructuras de roles y volumen de demanda.
4. `github_repos_real.csv`: Cuantificación de competencias técnicas y lenguajes.

Los perfiles tecnológicos han sido clusterizados en tres grandes *Macro-Categorías*:
* **Data & AI:** Data Scientists, ML Engineers, AI Engineers, Data Analysts.
* **Software Dev:** Frontend, Backend, Full-Stack, Mobile Developers.
* **Infra & Ops:** Cloud Architects, DevOps, Ciberseguridad.

## ⚙️ Despliegue en Vercel
Este repositorio está estructurado como una Single Page Application (SPA) estática. Para desplegarlo en Vercel, no se requieren configuraciones de Build o procesos de compilación Node.js.
