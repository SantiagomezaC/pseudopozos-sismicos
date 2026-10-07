# Pseudopozos sísmicos

[![Prueba reproducible](https://github.com/SantiagomezaC/pseudopozos-sismicos/actions/workflows/pruebas.yml/badge.svg)](https://github.com/SantiagomezaC/pseudopozos-sismicos/actions/workflows/pruebas.yml)
[![Versión](https://img.shields.io/github/v/release/SantiagomezaC/pseudopozos-sismicos?label=versi%C3%B3n)](https://github.com/SantiagomezaC/pseudopozos-sismicos/releases/latest)
[![Python 3.12+](https://img.shields.io/badge/python-3.12%2B-3776AB?logo=python&logoColor=white)](https://www.python.org)
[![Windows 10/11](https://img.shields.io/badge/plataforma-Windows%2010%2F11-0078D6?logo=windows&logoColor=white)](#requisitos)
[![Licencia MIT](https://img.shields.io/badge/licencia-MIT-green)](LICENSE)

Herramienta de escritorio para generar **pseudopozos georreferenciados** a partir de secciones sísmicas 2D interpretadas. Desarrollada como parte del trabajo de grado *Evaluación tectonoestratigráfica y cuantitativa del sistema fuente–sumidero de la cuenca Cesar–Ranchería* (Programa de Geología, Universidad del Norte).

A partir de la imagen de una sección, el programa identifica la línea en el catálogo de navegación sísmica del **Banco de Información Petrolera (BIP)** de la Agencia Nacional de Hidrocarburos (ANH), permite calibrar y digitalizar los horizontes interpretados y calcula, en puntos equiespaciados a lo largo de la línea, el espesor de las unidades cenozoicas comprendidas entre la superficie y el límite con el basamento. El resultado es una base de datos en Excel y CSV que QGIS y ArcGIS cargan directamente como capa de puntos.

## Contenido

- [Descarga rápida](#descarga-rápida)
- [Requisitos](#requisitos)
- [Instalación desde el código](#instalación-desde-el-código)
- [Método](#método)
- [Uso](#uso)
- [Resultados](#resultados)
- [Verificación](#verificación)
- [Limitaciones](#limitaciones)
- [Estructura del repositorio](#estructura-del-repositorio)
- [Cómo citar](#cómo-citar)
- [Licencia y fuente de datos](#licencia-y-fuente-de-datos)

## Descarga rápida

Quien solo quiera usar el programa, sin instalar Python, puede descargar el ejecutable de la [última versión publicada](https://github.com/SantiagomezaC/pseudopozos-sismicos/releases/latest). Basta con descomprimir el archivo `.zip` y abrir `Pseudopozos.exe`. El archivo debe permanecer dentro de su carpeta, junto a las librerías que lo acompañan.

## Requisitos

El programa requiere Windows 10 u 11, porque la identificación automática de la línea usa el motor de reconocimiento de texto (OCR) integrado en el sistema operativo. Para ejecutarlo desde el código se necesita, además, Python 3.12 o superior de [python.org](https://www.python.org), instalado con la opción *Add python.exe to PATH*. Se verificó con Python 3.14.7 en Windows 11.

## Instalación desde el código

1. Descargue el repositorio con **Code → Download ZIP** y descomprímalo, o clónelo:
   ```
   git clone https://github.com/SantiagomezaC/pseudopozos-sismicos.git
   ```
2. Haga doble clic en `instalar.bat`. El archivo crea un entorno virtual `.venv`, instala las versiones exactas de `requirements.txt` y ejecuta la prueba reproducible. La instalación es correcta si termina con *22 de 22 comprobaciones correctas*.
3. Abra el programa con `ejecutar.bat`.

Los mismos pasos desde la línea de comandos:

```
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python pruebas\prueba_reproducible.py
.venv\Scripts\python pseudopozos_sismicos.py
```

## Método

**Identificación de la línea.** El texto impreso en la imagen se lee con el OCR de Windows a tres escalas (×2, ×3 y ×4), debido a que los rótulos de las secciones suelen ser pequeños, y se combina con el nombre del archivo. Cada texto y cada nombre del catálogo se descomponen en bloques de letras y números sin separadores ni ceros a la izquierda, de modo que `CV-1979-08`, `cv 1979 8` y `CV_1979_08` resultan equivalentes; además se corrigen las confusiones habituales del OCR junto a dígitos (I o L leídas como 1, O como 0), aplicando la misma corrección al catálogo para que la comparación sea simétrica. Un nombre del catálogo solo se acepta si todos sus bloques aparecen consecutivos en un mismo texto, lo que impide confundir `CV-1979-08` con `CV-1979-080`. Cuando un nombre está contenido en otro (el programa `CV-1979` dentro de la línea `CV-1979-08`) se conserva el más específico. El usuario confirma la identificación a la vista de los atributos de la línea en el BIP.

**Calibración y digitalización.** Dos clics fijan los extremos de la sección y otros dos, el nivel de dos referencias verticales (por ejemplo 0,5 y 1,0 s TWT). Con ellos se definen dos transformaciones lineales independientes, de píxel a distancia a lo largo de la línea y de píxel a tiempo doble o profundidad. El usuario indica cuántas unidades tiene la línea y cuáles son, y digitaliza la superficie (opcional), la base de cada unidad salvo la más profunda y el límite con el basamento, que corresponde a la base de la unidad más profunda presente. Una unidad que se acuña se digitaliza solo donde existe; fuera de su horizonte su espesor es nulo. Los vértices marcados a menos de 10 píxeles de otro horizonte se proyectan sobre él, de manera que los acuñamientos coinciden exactamente y no generan espesores residuales.

**Cálculo de espesores.** En cada pseudopozo se leen los horizontes por interpolación lineal y se recorre la columna de techo a base. El espesor de cada unidad es

$$e_i = \frac{V_i\,(TWT_{base,i} - TWT_{tope,i})}{2},$$

donde la división por dos convierte el tiempo doble de viaje en tiempo sencillo. Las velocidades interválicas por defecto son 1150 m/s para el Cuaternario (capa 1) y 2438,4 m/s para el Neógeno (capa 2) y el Paleógeno (capa 3); todas son editables. Sin superficie digitalizada, el tope de la columna es TWT = 0. En las secciones convertidas a profundidad el espesor total es la diferencia directa de profundidades y las columnas por unidad se reportan como `NA`.

**Georreferenciación.** La posición de cada pseudopozo se expresa como fracción de la longitud de la sección y se lleva a la misma fracción de la longitud de la línea BIP, de forma que el resultado no depende de la escala gráfica de la imagen. La interpolación se realiza entre los vértices de la línea en su sistema de coordenadas original (MAGNA-SIRGAS Bogotá, EPSG:3116, en el catálogo de la ANH) y solo después se transforma a WGS84 (EPSG:4326) y a MAGNA-SIRGAS Origen Nacional (EPSG:9377). Si la línea se reproyectara antes de interpolar, sus segmentos rectos dejarían de serlo y los puntos se separarían ligeramente de la geometría oficial. Para cada punto se calcula su distancia a la línea BIP, que se exporta en la columna `Desv_BIP_m` y, en las pruebas, no supera 10⁻¹⁰ m.

## Uso

El catálogo BIP se configura una sola vez, ya sea descargando la capa *Sísmica 2D* del geovisor de la ANH con el botón correspondiente o cargando la navegación descargada del BIP (shapefile, GeoPackage, KML o geodatabase). El programa la recuerda en `%APPDATA%\Pseudopozos`. A partir de ahí el flujo sigue los pasos numerados del panel:

1. Cargar la imagen de la sección; la línea se identifica y se orienta.
2. Marcar los extremos de la sección.
3. Marcar dos referencias verticales.
4. Definir las unidades presentes y sus velocidades.
5. Digitalizar los horizontes.
6. Fijar la cantidad o el espaciado de los pseudopozos.
7. Calcular y exportar.

El botón *Ayuda* del programa describe cada paso.

## Resultados

Cada exportación produce un Excel con tres hojas, un CSV, una imagen de control con los horizontes y los pseudopozos dibujados sobre la sección y un archivo de proyecto `.json` que permite retomar o auditar la digitalización.

| Hoja | Contenido |
|---|---|
| `Pseudopozos` | Punto de control, pseudopozo, línea, longitud, latitud, TWT de cada base, velocidades, espesores por unidad y total |
| `GIS` | Los mismos datos con nombres de campo sin tildes ni espacios, coordenadas en EPSG:4326, EPSG:9377 y en el sistema original del BIP, y desviación respecto a la línea |
| `Metadatos` | Por línea: forma de identificación, catálogo y sistema de referencia usados, velocidades y desviación máxima |

Si se elige un Excel existente, la línea se agrega a la base y la numeración de los puntos de control continúa. En QGIS el CSV se carga con *Capa → Añadir capa → Añadir capa de texto delimitado*, tomando `Longitud` como X, `Latitud` como Y y EPSG:4326. En ArcGIS Pro se usa la herramienta *XY Table To Point* sobre la hoja `GIS` o el CSV.

## Verificación

[`pruebas/prueba_reproducible.py`](pruebas/prueba_reproducible.py) ejecuta 22 comprobaciones con datos sintéticos generados por la propia prueba, y GitHub Actions las repite en un Windows limpio en cada cambio del código (insignia al inicio de esta página). La más directa reproduce las cinco primeras filas de la base de puntos de control de la línea CV-1989-950: con TWT base Q = 0,47 s, TWT base N = 0,626 s y V_Q = 1500 m/s, el programa obtiene 352,5 m para el Cuaternario, 190,1952 m para el Neógeno y 542,6952 m de espesor total, idénticos a los de la tabla original. Las demás comprobaciones cubren:

- unidades ausentes y acuñadas, y secciones en profundidad;
- identificación de nombres con errores típicos de OCR y con distractores;
- unión de tramos de una misma línea y orientación;
- exactitud geoespacial y exportación.

## Limitaciones

La exactitud perpendicular a la línea es la de la geometría del BIP; la incertidumbre que permanece es longitudinal y depende de que los extremos marcados en la imagen correspondan a los extremos de la línea BIP. Cuando la imagen muestra solo un tramo, debe desactivarse el ajuste a la longitud BIP e indicarse el kilómetro de la línea en que comienza la sección. El OCR puede no leer rótulos muy pequeños o superpuestos; en ese caso la orientación se pregunta al usuario y la línea puede buscarse manualmente por su nombre. La descarga del catálogo depende de la disponibilidad del servicio público de la ANH; si no responde, puede usarse la navegación descargada del BIP. Finalmente, la conversión tiempo-profundidad usa una velocidad interválica constante por unidad, por lo que no representa variaciones laterales de velocidad.

## Estructura del repositorio

```
pseudopozos-sismicos/
├── pseudopozos_sismicos.py        Programa completo, documentado sección por sección
├── requirements.txt               Versiones exactas de las dependencias
├── instalar.bat                   Instalación reproducible en un entorno aislado
├── ejecutar.bat                   Abre el programa
├── pruebas/
│   └── prueba_reproducible.py     Verificación automática (22 comprobaciones)
├── docs/
│   └── Anexo_Codigo_Pseudopozos.docx   Código fuente en formato de anexo de tesis
├── CITATION.cff                   Datos para citar el software
└── LICENSE                        Licencia MIT
```

## Cómo citar

Si utiliza esta herramienta, cítela como:

> Meza Castro, M. S. (2026). *Pseudopozos sísmicos* (Versión 2.0) [Software]. Universidad del Norte. https://github.com/SantiagomezaC/pseudopozos-sismicos

GitHub genera la cita en formato APA o BibTeX desde el botón **Cite this repository**, a partir de [`CITATION.cff`](CITATION.cff).

## Licencia y fuente de datos

El código se distribuye bajo la [licencia MIT](LICENSE). La navegación sísmica proviene del Banco de Información Petrolera (BIP) y del geovisor de la Agencia Nacional de Hidrocarburos (ANH), capa *Sísmica 2D* (`https://geovisor.anh.gov.co/server/rest/services/GEOVISOR_v32/ANH_InsGDB/MapServer/2`); este repositorio no redistribuye esos datos.
