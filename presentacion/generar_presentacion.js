const pptx = require("pptxgenjs");
const fs = require("fs");

const P = new pptx();
P.layout = "LAYOUT_WIDE";            // 13.333 x 7.5
P.author = "Andrés Acuña";
P.title  = "Madera comercial en Manabí";

// Paleta tomada del tema de la plantilla
const NAVY = "0E2841", BLOQUE = "002060", AZUL = "156082",
      NARANJA = "E97132", VERDE = "196B24", CELESTE = "0F9ED5",
      GRIS = "5A6B7B", CLARO = "E8E8E8", BLANCO = "FFFFFF";

const PIE = "Proyecto de Vinculación: Práctica de Servicio Comunitario a la ESPOCH – GRUPO A";
const LOGO = "img/logo.png";
const SZ = JSON.parse(fs.readFileSync("img/sizes.json", "utf8"));

const notas = [];

// Encaja la imagen dentro de una caja conservando su proporción
function encajar(f, bx, by, bw, bh) {
  const [w, h] = SZ["img/" + f];
  const r = w / h, rb = bw / bh;
  let W, H;
  if (r > rb) { W = bw; H = bw / r; } else { H = bh; W = bh * r; }
  return { x: bx + (bw - W) / 2, y: by + (bh - H) / 2, w: W, h: H };
}

function cromo(s, titulo) {
  s.addShape(P.ShapeType.rect, { x: 0, y: 0.25, w: 6.6, h: 0.72, fill: { color: BLOQUE } });
  s.addText(titulo, { x: 0.22, y: 0.25, w: 6.2, h: 0.72, isTextBox: true,
    fontSize: 22, bold: true, color: BLANCO, fontFace: "Calibri", valign: "middle", margin: 0 });
  s.addImage({ path: LOGO, x: 0.47, y: 6.42, w: 0.85, h: 0.76 });
  s.addText(PIE, { x: 8.6, y: 6.68, w: 4.5, h: 0.6, isTextBox: true,
    fontSize: 9, color: GRIS, fontFace: "Calibri", align: "right", margin: 0 });
}

// Diapositiva de contenido: figura grande + lectura breve
function figura(titulo, archivo, bajada, lectura, nota) {
  const s = P.addSlide();
  s.background = { color: BLANCO };
  cromo(s, titulo);
  if (bajada) s.addText(bajada, { x: 0.5, y: 1.06, w: 12.3, h: 0.34, isTextBox: true,
    fontSize: 13, color: GRIS, fontFace: "Calibri", margin: 0 });

  const caja = encajar(archivo, 0.5, 1.45, 8.5, 4.75);
  s.addImage(Object.assign({ path: "img/" + archivo }, caja));

  s.addShape(P.ShapeType.roundRect, { x: 9.25, y: 1.5, w: 3.6, h: 4.6,
    fill: { color: "F4F6F8" }, rectRadius: 0.08 });
  s.addText("LECTURA", { x: 9.5, y: 1.7, w: 3.1, h: 0.26, isTextBox: true,
    fontSize: 10, bold: true, color: AZUL, fontFace: "Calibri", charSpacing: 1.5, margin: 0 });
  s.addText(lectura.map((t, i) => ({ text: t,
    options: { bullet: true, breakLine: i < lectura.length - 1, paraSpaceAfter: 8 } })),
    { x: 9.5, y: 2.05, w: 3.1, h: 3.9, isTextBox: true, fontSize: 12.5,
      color: NAVY, fontFace: "Calibri", margin: 0 });

  s.addNotes(nota);
  notas.push({ t: titulo, n: nota });
  return s;
}

// ── 1. Portada ────────────────────────────────────────────────────────────
{
  const s = P.addSlide();
  s.background = { color: NAVY };
  s.addText("¿Cuánta madera habrá en Manabí?", { x: 0.9, y: 2.0, w: 11.5, h: 1.1,
    isTextBox: true, fontSize: 42, bold: true, color: BLANCO, fontFace: "Calibri", margin: 0 });
  s.addText("Disponibilidad actual y futura de madera comercial bajo escenarios de cambio climático",
    { x: 0.9, y: 3.15, w: 10.5, h: 0.8, isTextBox: true, fontSize: 18, color: CLARO,
      fontFace: "Calibri", margin: 0 });
  s.addText("Andrés Acuña", { x: 0.9, y: 4.6, w: 6, h: 0.4, isTextBox: true,
    fontSize: 15, bold: true, color: BLANCO, fontFace: "Calibri", margin: 0 });
  s.addText("Modelación de nicho climático · CMIP6 · Inventario forestal nacional",
    { x: 0.9, y: 5.0, w: 8, h: 0.4, isTextBox: true, fontSize: 12, color: CELESTE,
      fontFace: "Calibri", margin: 0 });
  s.addImage({ path: LOGO, x: 11.5, y: 5.9, w: 1.0, h: 0.89 });
  s.addText(PIE, { x: 0.9, y: 6.6, w: 6.5, h: 0.5, isTextBox: true, fontSize: 9,
    color: GRIS, fontFace: "Calibri", margin: 0 });
  const n = "Buenos días. Voy a contarles un trabajo que empezó con una pregunta simple: cuánta madera comercial hay en Manabí, y cuánta habrá dentro de cincuenta años si el clima cambia. La pregunta parece de inventario forestal, pero terminó siendo un problema de datos y de modelación. Les adelanto el final: sí llegamos a metros cúbicos, pero con supuestos que voy a declarar en voz alta, porque esa es la diferencia entre un número útil y un número bonito.";
  s.addNotes(n); notas.push({ t: "Portada", n });
}

// ── 2. La pregunta y el obstáculo ─────────────────────────────────────────
{
  const s = P.addSlide();
  s.background = { color: BLANCO };
  cromo(s, "El problema de partida");
  const tarjetas = [
    { t: "Lo que se quería", d: "Metros cúbicos de madera comercial, hoy y en 2070, por especie y por cantón.", c: AZUL },
    { t: "Lo que hay", d: "El inventario nacional publicado registra un árbol testigo por especie y parcela, no todos los fustes.", c: NARANJA },
    { t: "La consecuencia", d: "Sin densidad de árboles por hectárea no se puede expandir a superficie. El atajo daría cifras infladas.", c: "B04A3A" }
  ];
  tarjetas.forEach((k, i) => {
    const x = 0.7 + i * 4.15;
    s.addShape(P.ShapeType.roundRect, { x, y: 1.6, w: 3.85, h: 2.7,
      fill: { color: "F4F6F8" }, rectRadius: 0.1 });
    s.addShape(P.ShapeType.ellipse, { x: x + 0.28, y: 1.9, w: 0.5, h: 0.5, fill: { color: k.c } });
    s.addText(String(i + 1), { x: x + 0.28, y: 1.9, w: 0.5, h: 0.5, isTextBox: true,
      fontSize: 17, bold: true, color: BLANCO, align: "center", valign: "middle",
      fontFace: "Calibri", margin: 0 });
    s.addText(k.t, { x: x + 0.28, y: 2.55, w: 3.3, h: 0.4, isTextBox: true,
      fontSize: 16, bold: true, color: NAVY, fontFace: "Calibri", margin: 0 });
    s.addText(k.d, { x: x + 0.28, y: 2.98, w: 3.3, h: 1.2, isTextBox: true,
      fontSize: 12.5, color: GRIS, fontFace: "Calibri", margin: 0 });
  });
  s.addShape(P.ShapeType.roundRect, { x: 0.7, y: 4.65, w: 12.15, h: 1.35,
    fill: { color: NAVY }, rectRadius: 0.1 });
  s.addText("La salida no fue abandonar la pregunta, sino declarar el supuesto que falta y medir cuánto pesa.",
    { x: 1.05, y: 4.95, w: 11.5, h: 0.8, isTextBox: true, fontSize: 17, bold: true,
      color: BLANCO, fontFace: "Calibri", margin: 0 });
  const n = "Aquí está el nudo del trabajo. Queríamos metros cúbicos. Pero al abrir el inventario nacional publicado en GBIF descubrimos que no es un censo: registra aproximadamente un árbol testigo por especie y por parcela. Lo comprobamos de tres formas — la densidad implícita daba diecinueve árboles por hectárea cuando un bosque tropical tiene entre cuatrocientos y seiscientos; ninguna especie se repetía dentro de una parcela; y la numeración de campo tenía huecos. Sin densidad no hay expansión a hectárea. La tentación era decir 'no se puede' y entregar solo el análisis climático. Decidimos lo contrario: declarar el supuesto que falta, ponerle un rango, y medir cuánto mueve el resultado.";
  s.addNotes(n); notas.push({ t: "El problema de partida", n });
}

// ── 3. Método en tres piezas ──────────────────────────────────────────────
{
  const s = P.addSlide();
  s.background = { color: BLANCO };
  cromo(s, "Cómo lo resolvimos");
  const pasos = [
    { t: "Dónde puede crecer", d: "Modelo de nicho climático por especie, con clima actual y proyectado a 2070.", c: VERDE },
    { t: "Dónde hay bosque", d: "Cobertura arbórea satelital (Hansen, 30 m). Filtra pasto, cultivo y ciudad.", c: AZUL },
    { t: "Cuánto rinde", d: "Volumen por árbol medido en el inventario, por densidad de fustes supuesta.", c: NARANJA }
  ];
  pasos.forEach((k, i) => {
    const y = 1.55 + i * 1.6;
    s.addShape(P.ShapeType.ellipse, { x: 0.8, y: y + 0.12, w: 0.85, h: 0.85, fill: { color: k.c } });
    s.addText(String(i + 1), { x: 0.8, y: y + 0.12, w: 0.85, h: 0.85, isTextBox: true,
      fontSize: 26, bold: true, color: BLANCO, align: "center", valign: "middle",
      fontFace: "Calibri", margin: 0 });
    s.addText(k.t, { x: 2.0, y: y + 0.1, w: 4.2, h: 0.45, isTextBox: true,
      fontSize: 19, bold: true, color: NAVY, fontFace: "Calibri", margin: 0 });
    s.addText(k.d, { x: 2.0, y: y + 0.56, w: 6.4, h: 0.6, isTextBox: true,
      fontSize: 13.5, color: GRIS, fontFace: "Calibri", margin: 0 });
  });
  s.addShape(P.ShapeType.roundRect, { x: 8.9, y: 1.55, w: 3.95, h: 4.45,
    fill: { color: "F4F6F8" }, rectRadius: 0.1 });
  s.addText("Las tres son necesarias", { x: 9.2, y: 1.9, w: 3.4, h: 0.5, isTextBox: true,
    fontSize: 16, bold: true, color: NAVY, fontFace: "Calibri", margin: 0 });
  s.addText([
    { text: "Clima sin cobertura: dice dónde podría, no dónde hay.", options: { bullet: true, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Cobertura sin rendimiento: da hectáreas, no metros cúbicos.", options: { bullet: true, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Rendimiento sin clima: no proyecta a futuro.", options: { bullet: true, paraSpaceAfter: 10 } }
  ], { x: 9.2, y: 2.5, w: 3.4, h: 3.2, isTextBox: true, fontSize: 13, color: GRIS,
       fontFace: "Calibri", margin: 0 });
  const n = "El método tiene tres piezas y las tres hacen falta. La primera: un modelo de nicho climático que aprende qué clima tolera cada especie y lo proyecta a 2070. La segunda: cobertura arbórea satelital, que nos dice dónde hay árboles de verdad y filtra el pasto, el cultivo y la ciudad. La tercera: cuánto rinde una hectárea, que sale del volumen por árbol medido en campo multiplicado por una densidad supuesta. Si quitan cualquiera de las tres el resultado deja de tener sentido. Con solo clima dirían dónde podría crecer, no dónde hay. Con solo cobertura tendrían hectáreas, no metros cúbicos. Con solo rendimiento no habría futuro.";
  s.addNotes(n); notas.push({ t: "Cómo lo resolvimos", n });
}

// ── 4 a 17. Diapositivas de figura ────────────────────────────────────────
figura("El área de estudio", "fig00.png",
  "Manabí: 22 cantones, 19.603 km²",
  ["Los resultados se agregan por cantón, que es la unidad de gestión.",
   "La provincia combina bosque seco costero con plantaciones comerciales.",
   "El modelo se calibra con un halo de 150 km alrededor para captar el gradiente climático."],
  "Manabí tiene veintidós cantones y casi veinte mil kilómetros cuadrados. Elegimos el cantón como unidad de agregación porque es donde se toman las decisiones de manejo. Un detalle metodológico: el modelo no se entrena solo con Manabí, sino con un halo de ciento cincuenta kilómetros alrededor. Si entrenáramos solo dentro de la provincia el modelo no vería suficiente variación climática para aprender qué prefiere cada especie.");

figura("De dónde aprende el modelo", "fig01.png",
  "Registros de presencia de las cinco especies seleccionadas",
  ["Cada punto es un lugar donde alguien documentó la especie.",
   "El modelo no ve el mapa: ve el clima de esos puntos.",
   "La concentración cerca de vías y poblados es un sesgo conocido que declaramos."],
  "Esto es lo que alimenta al modelo. Cada punto es un registro verificado de presencia, descargado de GBIF. Es importante entender que el modelo no ve el mapa: ve el clima de esos puntos. Y hay un sesgo que conviene declarar antes de que alguien lo pregunte: los registros se concentran cerca de carreteras y poblados, porque ahí es donde la gente observa. Eso significa que el modelo aprende en parte dónde hay observadores, no solo dónde hay especies. No lo corregimos; lo declaramos.");

figura("Cuánto cambia el clima", "fig02.png",
  "Aumento de temperatura proyectado, promedio provincial",
  ["Es el motor de todo: si el clima no cambiara, la idoneidad tampoco.",
   "La diferencia entre escenarios se abre con el tiempo.",
   "SSP1-2.6 es el escenario optimista; SSP5-8.5 el de altas emisiones."],
  "Antes de hablar de especies conviene ver el motor. Esto es cuánto sube la temperatura media en Manabí según los tres escenarios. Fíjense en dos cosas. Primero, sube en los tres, incluso en el optimista: eso ya está comprometido. Segundo, la diferencia entre escenarios se abre con el tiempo — hacia 2030 son parecidos, hacia 2070 no. Eso significa que las decisiones de emisiones de esta década determinan el escenario de finales de siglo, no el de la próxima.");

figura("Cómo funciona el modelo", "fig08.png",
  "Curvas de respuesta: qué clima prefiere cada especie",
  ["Variable respuesta: 1 donde hay registro, 0 en puntos de fondo.",
   "El modelo aprende qué combinaciones de clima separan unos de otros.",
   "Curva que sube: más de esa variable favorece. Campana: hay un óptimo.",
   "Línea plana: el modelo no usa esa variable."],
  "Aquí abrimos la caja negra. La variable que el modelo predice es binaria: vale uno donde hay un registro de la especie y cero en puntos de fondo tomados al azar en la región. El modelo aprende qué combinaciones de clima separan unos de otros. Estas curvas muestran qué predice al mover una variable dejando las demás en su valor mediano. Si la curva sube, más de esa variable favorece a la especie. Si hace una campana, hay un óptimo intermedio. Si sale plana, el modelo simplemente no usa esa variable. Los escalones bruscos son la firma de los modelos de árboles, que parten el rango en tramos.");

figura("¿Y funciona?", "fig05.png",
  "Desempeño de los tres algoritmos, validado en zonas no vistas",
  ["Comparamos Random Forest, XGBoost y GAM.",
   "La validación es espacial: el modelo se evalúa en zonas geográficas completas que no vio.",
   "Solo dos especies superan el umbral aceptable.",
   "Las otras tres se reportan, pero no sostienen conclusiones."],
  "Esta es la diapositiva más honesta de la presentación. Comparamos tres algoritmos y los validamos de forma espacial: el modelo se entrena en unas zonas geográficas y se evalúa en otras que no vio nunca. Si partiéramos al azar, puntos vecinos quedarían a ambos lados y el modelo sacaría buenas notas por reconocer sitios ya vistos. El resultado: solo dos de las cinco especies superan el umbral aceptable. Las otras tres tienen un AUC cercano a 0,60, apenas mejor que lanzar una moneda. Las reportamos porque ocultarlas sería peor, pero sus proyecciones no sostienen ninguna conclusión por precisa que parezca la cifra.");

figura("Dónde puede crecer hoy", "fig06.png",
  "Idoneidad climática actual, una especie por panel",
  ["Verde intenso: el clima se parece al de los sitios donde la especie vive.",
   "No dice que haya árboles ahí, ni que vayan a llegar.",
   "Dice que el clima no sería el impedimento."],
  "Este es el resultado del modelo para el presente. Verde intenso significa que el clima de ese píxel se parece al de los lugares donde la especie está documentada. Y aquí hay que ser muy preciso con el lenguaje, porque es donde más se confunde la gente: el mapa no dice que haya árboles ahí, ni que vayan a llegar. Dice que el clima no sería el impedimento. Entre 'el clima lo permite' y 'ahí hay un bosque' median la dispersión de semillas, el uso del suelo y la demografía del rodal, y ninguna de las tres está en el modelo.");

figura("Qué variable manda", "fig07.png",
  "Importancia de cada variable climática",
  ["Barra larga: el modelo depende de esa variable para decidir.",
   "Dominan las variables de lluvia sobre las de temperatura.",
   "Es coherente con bosque seco: el factor limitante es el agua."],
  "Esta figura responde a qué mira el modelo para decidir. Mide cuánto empeora la predicción si se desordena una variable. Lo interesante es que dominan las variables de lluvia por encima de las de temperatura. Eso es coherente con lo que sabemos del bosque seco tumbesino: ahí el factor limitante es el agua, no el calor. Y tiene una implicación práctica — si el cambio climático altera el régimen de lluvias más que la temperatura, el efecto sobre estas especies será mayor de lo que sugeriría mirar solo el calentamiento.");

figura("El panorama a 2070", "fig11.png",
  "Cambio de idoneidad por especie, escenario y horizonte",
  ["Rojo: pierde aptitud climática. Azul: gana.",
   "Si el color se intensifica hacia la derecha, el efecto crece con el tiempo.",
   "Si se intensifica hacia abajo, depende del escenario de emisiones."],
  "Si tuviera que quedarme con una sola figura sería esta. Cada panel es una especie, cada fila un escenario de emisiones y cada columna un horizonte temporal. Rojo es pérdida de aptitud, azul es ganancia. Léanla en dos direcciones: si el color se intensifica hacia la derecha, el efecto crece con el tiempo; si se intensifica hacia abajo, depende del escenario de emisiones. El caso más llamativo es Cordia alliodora, el laurel, que pierde más de la mitad de su idoneidad bajo altas emisiones. Es además una de las especies maderables más usadas en la costa.");

figura("El mapa del cambio", "fig10.png",
  "Diferencia de idoneidad entre hoy y 2070, escenario de altas emisiones",
  ["Rojo: el clima se aleja del que la especie tolera.",
   "Azul: se acerca.",
   "Plantar hoy en zona roja es apostar contra la tendencia."],
  "El mismo resultado, ahora en el espacio. Rojo significa que hacia 2070 el clima de ese lugar se aleja del que la especie tolera; azul, que se acerca. La lectura operativa es directa: si están planificando una plantación a turno largo, plantar hoy en una zona roja es apostar contra la tendencia. Y conviene recordar algo que dije antes — la pérdida de idoneidad no se manifiesta primero como árboles muertos, sino como plántulas que no prosperan. El árbol adulto aguanta décadas; el problema aparece en la regeneración.");

figura("Dónde plantar qué", "fig13.png",
  "Los cinco cantones más aptos para cada especie",
  ["Es la tabla de decisión si hay presupuesto para plantar.",
   "Los valores absolutos son bajos porque se comparan contra toda la región.",
   "Lo que importa es el orden, no la cifra."],
  "Esta es la diapositiva accionable. Si mañana hay presupuesto para un programa de plantación, aquí está a qué cantón ir con qué especie. Una advertencia sobre los números: los valores absolutos son bajos porque el modelo compara cada sitio contra todo el fondo regional, no contra un ideal. Lo que importa es el orden, no la cifra. Y noten que los cantones no se repiten entre especies, lo cual tiene sentido: cada una tiene su propio nicho. Por eso, dicho sea de paso, no promediamos la idoneidad entre especies — mezclar nichos distintos da un número sin significado.");

figura("Cuánto bosque hay realmente", "fig17.png",
  "Cobertura arbórea de Manabí según Hansen, 30 m de resolución",
  ["70,1% de la provincia tiene dosel sobre el 30%.",
   "No es un resultado: es un insumo del cálculo de volumen.",
   "Es la capa que separa 'donde podría haber' de 'donde hay'."],
  "Esta no es una figura de resultados, es un insumo, y por eso la explico aparte. Es cobertura arbórea medida por satélite a treinta metros de resolución. Nos dice que el setenta por ciento de Manabí tiene dosel por encima del treinta por ciento. Su papel en el trabajo es el de filtro: el mapa de idoneidad dice dónde el clima permitiría una especie, y este dice dónde hay árboles de hecho. Solo la intersección de los dos produce metros cúbicos creíbles. La siguiente diapositiva muestra por qué eso importa tanto.");

figura("El atajo que infla la cifra", "fig18.png",
  "Superficie idónea frente a superficie idónea con bosque",
  ["Barra gris: lo que el clima permite.",
   "Barra verde: la parte que además tiene bosque.",
   "La brecha es el error de saltarse el cruce con cobertura."],
  "Aquí está la trampa metodológica más fácil de caer, y quiero mostrarla explícitamente porque la vi propuesta en más de un lado. El atajo sería multiplicar el rendimiento por hectárea por la superficie climáticamente idónea. La barra gris es esa superficie; la verde es la parte que además tiene bosque. La distancia entre ambas es el error que se cometería. En nuestro caso la cifra saldría varias veces inflada. Y hay un segundo problema que no se ve en el gráfico: las áreas idóneas de las cinco especies se solapan, así que sumarlas contaría el mismo terreno varias veces. Por eso las presentamos como alternativas, nunca como un total provincial.");

figura("Cuántos metros cúbicos", "fig19.png",
  "Volumen potencial por especie, actual y proyectado",
  ["La línea es el valor central; la banda, el rango de densidad supuesta.",
   "La banda es ancha porque ese supuesto no está medido.",
   "Lo defendible es la pendiente, no la altura."],
  "Y aquí está finalmente la respuesta en metros cúbicos. Pero léanla con cuidado. La línea es el valor central y la banda cubre el rango de densidad de fustes que supusimos, entre trescientos y seiscientos árboles por hectárea. La banda es ancha porque ese supuesto no está medido, y es honesto que se vea. Lo importante es esto: lo defendible de este gráfico es la pendiente, no la altura. El porcentaje de cambio entre hoy y 2070 no depende del supuesto de densidad, porque se cancela al hacer el cociente. El valor absoluto sí depende. Si alguien les pregunta la cifra exacta, esa es la respuesta correcta.");

figura("De qué depende el resultado", "fig20.png",
  "Sensibilidad: cuánto mueve cada supuesto",
  ["Cada escenario se recalcula, no se multiplica por un factor.",
   "La densidad de fustes entra como factor directo: su efecto es proporcional.",
   "Una sola campaña de medición en campo reduciría más la incertidumbre que cualquier refinamiento del modelo."],
  "Terminamos con la pregunta que siempre hace un evaluador: de qué depende su número. Este análisis varía un parámetro a la vez y recalcula todo desde el mapa — no multiplica el resultado por un factor, que sería suponer lo que queremos averiguar. La lección práctica es clara: la densidad de fustes entra como factor multiplicativo directo, así que su efecto es exactamente proporcional. Eso significa que una sola campaña de medición de densidad en campo reduciría la incertidumbre de todo el trabajo más que cualquier refinamiento del modelo climático. Ahí es donde hay que gastar el próximo dólar.");

// ── 18. Conclusiones ──────────────────────────────────────────────────────
{
  const s = P.addSlide();
  s.background = { color: NAVY };
  s.addText("Qué podemos afirmar", { x: 0.9, y: 0.75, w: 11.5, h: 0.75, isTextBox: true,
    fontSize: 34, bold: true, color: BLANCO, fontFace: "Calibri", margin: 0 });
  const filas = [
    { c: VERDE, t: "Sólido", d: "El cambio relativo de idoneidad y de volumen entre escenarios. Los supuestos se cancelan en el cociente." },
    { c: CELESTE, t: "Utilizable", d: "El ranking de cantones por especie, para decidir dónde plantar qué." },
    { c: NARANJA, t: "Con reservas", d: "El valor absoluto en metros cúbicos: es potencial condicionado, no existencias medidas." },
    { c: "B04A3A", t: "No afirmable", d: "Proyecciones de las tres especies con AUC bajo 0,70. Y un total provincial sumando especies." }
  ];
  filas.forEach((f, i) => {
    const y = 1.85 + i * 1.08;
    s.addShape(P.ShapeType.roundRect, { x: 0.9, y, w: 2.05, h: 0.62,
      fill: { color: f.c }, rectRadius: 0.08 });
    s.addText(f.t, { x: 0.9, y, w: 2.05, h: 0.62, isTextBox: true, fontSize: 14,
      bold: true, color: BLANCO, align: "center", valign: "middle", fontFace: "Calibri", margin: 0 });
    s.addText(f.d, { x: 3.2, y: y - 0.02, w: 9.2, h: 0.7, isTextBox: true, fontSize: 14,
      color: CLARO, fontFace: "Calibri", valign: "middle", margin: 0 });
  });
  s.addImage({ path: LOGO, x: 11.6, y: 6.3, w: 0.85, h: 0.76 });
  const n = "Cierro con lo que se puede afirmar y lo que no, que para mí es la parte más importante. Sólido: el cambio relativo entre escenarios, porque los supuestos se cancelan al hacer el cociente. Utilizable: el ranking de cantones por especie, que es directamente accionable. Con reservas: el valor absoluto en metros cúbicos, que es un potencial condicionado y no existencias medidas. Y no afirmable: las proyecciones de las tres especies cuyo modelo no supera el umbral, y cualquier total provincial que sume especies, porque sus áreas se solapan. Prefiero entregar cuatro afirmaciones defendibles que veinte que no resistan la primera pregunta.";
  s.addNotes(n); notas.push({ t: "Qué podemos afirmar", n });
}

// ── 19. Qué falta ─────────────────────────────────────────────────────────
{
  const s = P.addSlide();
  s.background = { color: BLANCO };
  cromo(s, "Qué haría falta para cerrar la brecha");
  const pasos = [
    { n: "1", t: "Inventario completo del MAATE", d: "Todos los fustes por parcela y los factores de expansión del muestreo. Convierte el supuesto de densidad en dato.", c: VERDE },
    { n: "2", t: "Parcelas permanentes con remedición", d: "Dos mediciones en el tiempo dan la tasa de crecimiento real, hoy supuesta.", c: AZUL },
    { n: "3", t: "Mapa de plantaciones del MAG", d: "Permite separar plantación de bosque natural, imposible con el dato actual.", c: NARANJA }
  ];
  pasos.forEach((k, i) => {
    const y = 1.5 + i * 1.55;
    s.addShape(P.ShapeType.roundRect, { x: 0.7, y, w: 12.15, h: 1.3,
      fill: { color: "F4F6F8" }, rectRadius: 0.1 });
    s.addShape(P.ShapeType.ellipse, { x: 1.0, y: y + 0.32, w: 0.66, h: 0.66, fill: { color: k.c } });
    s.addText(k.n, { x: 1.0, y: y + 0.32, w: 0.66, h: 0.66, isTextBox: true, fontSize: 20,
      bold: true, color: BLANCO, align: "center", valign: "middle", fontFace: "Calibri", margin: 0 });
    s.addText(k.t, { x: 2.0, y: y + 0.2, w: 10.4, h: 0.4, isTextBox: true, fontSize: 17,
      bold: true, color: NAVY, fontFace: "Calibri", margin: 0 });
    s.addText(k.d, { x: 2.0, y: y + 0.62, w: 10.4, h: 0.55, isTextBox: true, fontSize: 13,
      color: GRIS, fontFace: "Calibri", margin: 0 });
  });
  // A la derecha del logo, no encima: el logo ocupa x 0,47–1,32 / y 6,42–7,18
  s.addText("Gracias. ¿Preguntas?", { x: 1.65, y: 6.42, w: 5, h: 0.76, isTextBox: true,
    fontSize: 20, bold: true, color: NAVY, fontFace: "Calibri", valign: "middle", margin: 0 });
  const n = "Y termino con lo que falta, porque un trabajo honesto también dice qué le falta. Primero, el inventario completo del MAATE: tiene todos los fustes por parcela y los factores de expansión del muestreo, y convertiría nuestro supuesto de densidad en un dato medido. Segundo, parcelas permanentes con remedición: dos mediciones separadas en el tiempo dan la tasa de crecimiento real, que hoy suponemos. Tercero, el mapa de plantaciones del MAG, que permitiría separar plantación de bosque natural — algo imposible con el dato actual. Con esas tres piezas este trabajo deja de ser una estimación con supuestos y pasa a ser un inventario proyectado. Gracias, y quedo atento a preguntas.";
  s.addNotes(n); notas.push({ t: "Qué haría falta", n });
}

P.writeFile({ fileName: "Manabi_madera_presentacion.pptx" }).then(() => {
  // Guion del expositor
  const L = [];
  L.push("GUION DEL EXPOSITOR");
  L.push("Presentación: ¿Cuánta madera habrá en Manabí?");
  L.push("Andrés Acuña · " + PIE);
  L.push("Duración estimada: 18–22 minutos (" + notas.length + " diapositivas)");
  L.push("");
  L.push("=".repeat(78));
  notas.forEach((x, i) => {
    L.push("");
    L.push("DIAPOSITIVA " + (i + 1) + " — " + x.t);
    L.push("-".repeat(78));
    L.push("");
    // Párrafos de ~100 caracteres para leer cómodo
    let linea = "";
    x.n.split(" ").forEach(w => {
      if ((linea + " " + w).trim().length > 92) { L.push(linea.trim()); linea = w; }
      else linea += " " + w;
    });
    if (linea.trim()) L.push(linea.trim());
    L.push("");
  });
  L.push("=".repeat(78));
  L.push("");
  L.push("RECOMENDACIONES");
  L.push("");
  L.push("- Ritmo: aproximadamente un minuto por diapositiva de figura.");
  L.push("- Las diapositivas 9 y 17 son las que más preguntas generan: la de desempeño");
  L.push("  de los modelos y la del volumen. Conviene tenerlas bien ensayadas.");
  L.push("- Si el tiempo aprieta, se pueden omitir las diapositivas 11 (importancia de");
  L.push("  variables) y 14 (dónde plantar qué) sin romper el hilo.");
  L.push("- Pregunta probable: \"¿por qué no dan una cifra exacta?\". Respuesta: porque");
  L.push("  la densidad de fustes no está medida; el porcentaje de cambio sí es firme.");
  L.push("- Pregunta probable: \"¿esto dice dónde estarán los árboles?\". Respuesta: no,");
  L.push("  dice dónde el clima lo permitirá. Faltan dispersión, suelo y uso del suelo.");
  fs.writeFileSync("Guion_presentacion.txt", L.join("\n"), "utf8");
  console.log("listo:", notas.length, "diapositivas");
});
