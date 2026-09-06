# Base de datos de eficiencia de produccion

Planta transaccional de alimentos con 4 lineas (ketchup, queso, vinagre, mostaza),
3 turnos de 8 horas y 14 SKU. 3,756 registros de turno y 11,068 eventos de paro
entre el 1 de septiembre de 2025 y el 31 de agosto de 2026.

Los datos son simulados. La semilla y los parametros que los generan estan en la
hoja `Parametros` del libro de Excel y en `scripts/generar_base_datos.py`. No
provienen de ningun sistema ni de ninguna planta real.

Este repositorio no contiene el modelo de frontera eficiente ni ninguna rutina de
optimizacion. Termina en la matriz de retornos por periodo que produce
`BaseDatosPlanta.matriz_retornos()`.

## Estructura

```
.
├── app.py                       tablero de Streamlit
├── requirements.txt
├── .streamlit/config.toml       tema de la aplicacion
├── .github/workflows/
│   └── validacion.yml           corre las 37 pruebas en cada push
├── data/
│   ├── BD_Eficiencia_Planta.xlsx  fuente editable, con formulas vivas
│   └── csv/                       copia de lectura rapida para la app
├── scripts/
│   ├── generar_base_datos.py    reconstruye el xlsx desde cero
│   └── preparar_datos.py        convierte el xlsx a CSV
└── src/
    └── gestion_bd.py            carga, validacion, KPI y matriz de retornos
```

## Ejecutar en local

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

La aplicacion abre en http://localhost:8501. Lee `data/csv` si existe (0.04 s) y
cae al xlsx si no (alrededor de 5 s).

Sin interfaz grafica:

```bash
python src/gestion_bd.py --validar
python src/gestion_bd.py --resumen linea_turno
python src/gestion_bd.py --pareto 10
python src/gestion_bd.py --exportar-matriz matriz.csv --activo linea_turno --frecuencia D
```

## Publicar en GitHub

```bash
git init
git add .
git commit -m "Base de datos de eficiencia de produccion y tablero"
git branch -M main
git remote add origin https://github.com/USUARIO/REPOSITORIO.git
git push -u origin main
```

El repositorio pesa alrededor de 6 MB, muy por debajo del limite de 100 MB por
archivo de GitHub, asi que no hace falta Git LFS.

El flujo de `.github/workflows/validacion.yml` corre en cada push: instala las
dependencias, regenera los CSV desde el xlsx, ejecuta las 37 pruebas de
integridad y sube la matriz de retornos como artefacto descargable. Si una prueba
falla, el job falla, porque `--validar` devuelve codigo de salida 1.

## Publicar en Streamlit Community Cloud

1. Sube el repositorio a GitHub como publico.
2. Entra a https://share.streamlit.io con la misma cuenta de GitHub y autoriza el
   acceso al repositorio.
3. Elige el repositorio, la rama `main` y el archivo principal `app.py`.
4. En la configuracion avanzada selecciona Python 3.11.
5. Despliega. La primera compilacion instala `requirements.txt` y tarda unos
   minutos; las siguientes son mas rapidas.

Cada push a `main` vuelve a desplegar la aplicacion.

## Si regeneras el xlsx

`openpyxl` escribe las formulas sin valor en cache, asi que un libro recien
generado se lee vacio en las columnas calculadas. Despues de correr
`scripts/generar_base_datos.py` hay que abrir el archivo en Excel y guardarlo, o
recalcularlo con LibreOffice:

```bash
soffice --headless --convert-to xlsx --outdir data data/BD_Eficiencia_Planta.xlsx
python scripts/preparar_datos.py
```

`preparar_datos.py` se detiene con un mensaje explicito si detecta el libro sin
recalcular.

## Marco de yield

- Perfection cost = kg teoricos sin merma x precio de materia prima
- Costo de scrap estandar = perfection cost x factor de scrap del SKU
- Standard prime cost = perfection cost + costo de scrap estandar
- MUV = (kg reales - kg estandar) x precio de materia prima
- Act abs loss = costo de scrap estandar + MUV
- Yield loss to perfection = act abs loss / perfection cost

El rendimiento diario (kg de producto entre kg de materia prima) se reporta en
columna aparte porque no es la misma medida que el YTP.
