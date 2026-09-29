# Madera comercial en Manabí

Estimación de la disponibilidad actual y futura de madera comercial en la provincia
de Manabí (Ecuador), combinando modelación de nicho climático bajo escenarios CMIP6,
cobertura arbórea satelital e inventario forestal nacional.

Todo el análisis es reproducible y usa exclusivamente fuentes abiertas.

---

## Qué responde y qué no

Vale la pena fijar el alcance antes que los resultados, porque determina qué
afirmaciones son defendibles.

**Se estima con datos reales**

- Idoneidad climática por especie, en el presente y bajo tres escenarios de emisión
  en tres horizontes temporales.
- Superficie idónea por cantón y su variación proyectada.
- Volumen comercial individual, relación altura-diámetro e incremento proyectado, a
  partir de los árboles medidos por el inventario nacional.
- Incertidumbre entre algoritmos, entre escenarios y entre modelos climáticos.

**Se estima con supuestos declarados**

El volumen en metros cúbicos encadena superficie idónea, cobertura forestal real y
rendimiento por hectárea. El último término se apoya en un supuesto de densidad de
fustes, porque la publicación abierta del inventario nacional es un registro de
composición florística —un árbol testigo por especie y parcela— y no un censo. Ese
supuesto se declara, se presenta como rango y se somete a análisis de sensibilidad.

La consecuencia práctica: **el valor absoluto es un potencial condicionado, no
existencias medidas**, mientras que el **cambio relativo entre escenarios sí es
robusto**, porque los supuestos se cancelan en el cociente.

**No es estimable con el dato abierto**

- Existencias reales en pie verificadas en campo. Requiere el inventario completo del
  MAATE, con densidad por parcela y factores de expansión.
- Un total provincial sumando especies: las áreas idóneas se solapan, de modo que las
  cifras por especie son alternativas excluyentes, no componentes de un total.

---

## Estructura

```
analisis/
  manabi_madera.Rmd        Documento principal: nicho climático, CMIP6, volumen
  madera_comercial.Rmd     Dasometría del inventario nacional y serie mensual
  madera_satelital.Rmd     Estudio previo de factibilidad (resultado negativo)
presentacion/
  Manabi_madera_presentacion.pptx   19 diapositivas
  Guion_presentacion.txt            Qué decir en cada diapositiva
  generar_presentacion.js           Generador reproducible (pptxgenjs)
  figuras/                          Gráficos exportados del informe
```

---

## Cómo ejecutarlo

```r
# Desde analisis/
rmarkdown::render("manabi_madera.Rmd")
```

El documento crea su propia estructura de directorios (`data/raw`, `outputs/`,
`reports/`) y descarga lo que necesite. **Todo queda cacheado**: al volver a tejer no
se repite ninguna descarga. Una sección al inicio del informe reporta el estado de la
caché antes de tocar la red.

La primera ejecución descarga aproximadamente 8,5 GB de datos climáticos globales
(WorldClim y CMIP6 a 2,5 minutos de arco). Para una prueba rápida, reducir
`MODELOS_GCM` a un solo modelo en el bloque de parámetros.

### Dependencias

```r
install.packages(c("sf", "terra", "geodata", "dplyr", "tidyr", "ggplot2",
                   "scales", "ranger", "xgboost", "mgcv", "knitr", "jsonlite"))
```

Para regenerar la presentación: `node presentacion/generar_presentacion.js`
(requiere `pptxgenjs`).

---

## Fuentes de datos

| Fuente | Uso | Acceso |
|---|---|---|
| **GBIF** | Ocurrencias de especies e inventario forestal nacional (MAATE) | API pública, sin credenciales |
| **WorldClim 2.1** | Variables bioclimáticas actuales | Paquete `geodata` |
| **CMIP6 / WorldClim** | Proyecciones climáticas, SSP1-2.6 / 2-4.5 / 5-8.5 | Paquete `geodata` |
| **Hansen GFC v1.12** | Cobertura arbórea a 30 m | Lectura remota por `/vsicurl/` |
| **GADM** | Límites provinciales y cantonales | Paquete `geodata` |

---

## Decisiones metodológicas

Tres elecciones sostienen la validez de los resultados y conviene señalarlas:

**Validación cruzada espacial por bloques.** Las ocurrencias de GBIF se agrupan cerca
de vías y poblados. Una partición aleatoria dejaría puntos vecinos a ambos lados y
produciría métricas infladas. Se particiona en bloques de 0,5° para que el modelo se
evalúe sobre zonas geográficas que no vio.

**Umbral de idoneidad por especie.** Con miles de puntos de fondo frente a unos
cientos de presencias, el modelo devuelve probabilidades bajas en todo el dominio; un
corte fijo en 0,5 dejaría cero píxeles y borraría toda la señal. Cada especie usa el
umbral que maximiza sensibilidad más especificidad.

**Intersección con cobertura arbórea.** La idoneidad climática no es bosque.
Multiplicar rendimiento por hectárea por superficie idónea sin cruzar con cobertura
infla el resultado varias veces.

---

## Limitación principal

Los modelos describen **nicho climático realizado**, no potencial fisiológico ni
capacidad de dispersión. Que un sitio se vuelva climáticamente apto no significa que
la especie llegue ni que haya bosque allí. La pérdida de idoneidad se manifiesta
primero como falla de regeneración, no como mortalidad inmediata.

El horizonte más lejano disponible en el producto CMIP6 descargable es 2061-2080; no
existe 2081-2100 en esa fuente.

---

## Referencias principales

- Fick, S. E., & Hijmans, R. J. (2017). WorldClim 2. *International Journal of
  Climatology, 37*(12), 4302–4315.
- Hansen, M. C., et al. (2013). High-resolution global maps of 21st-century forest
  cover change. *Science, 342*(6160), 850–853.
- O'Neill, B. C., et al. (2016). The Scenario Model Intercomparison Project
  (ScenarioMIP) for CMIP6. *Geoscientific Model Development, 9*(9), 3461–3482.
- Roberts, D. R., et al. (2017). Cross-validation strategies for data with temporal,
  spatial, hierarchical, or phylogenetic structure. *Ecography, 40*(8), 913–929.

El informe incluye la lista completa en formato APA.

---

## Autor

**Andrés Acuña** — [@andrestorres23](https://github.com/andrestorres23)

Proyecto de Vinculación: Práctica de Servicio Comunitario a la ESPOCH.
