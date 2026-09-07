/* global document, Office, PowerPoint, FileReader, fetch */

const RUTA_MAESTRO = "/maestro.pptx";

/* ---------- Secciones y subsecciones ----------
   Reemplaza el sistema anterior de "Categoría visual" plana. Ahora cada
   lámina se clasifica en dos niveles: Sección (tema grande) y Subsección
   (dentro de esa sección). Todo es 100% manual, sin detección automática
   -- el campo tipo_plantilla de indexador.py se sigue mostrando como dato
   informativo aparte, pero ya no alimenta la clasificación. */

const SECCIONES_DEFAULT = {
  "Portada y cierre": ["Portada", "Índice / contenido", "Cierre"],
  "Quiénes somos": [
    "Ecosistema y footprint",
    "Por qué Matrix",
    "Equipo y CVs",
    "Track record",
  ],
  "Diagnóstico y contexto": [
    "Contexto del negocio",
    "Benchmarks y comparables",
    "Línea base",
  ],
  "Objetivos y enfoque": [
    "Objetivos",
    "Alcance y plan de trabajo",
    "Cronograma y olas",
  ],
  "Frameworks y metodologías": [
    "Frameworks conceptuales",
    "Metodologías de análisis",
    "Herramientas propietarias",
  ],
  "Análisis y hallazgos": ["Gráficos de datos", "Casos y ejemplos", "Mapas"],
  "Entregables e impacto": ["Entregables por fase", "Impactos y resultados"],
  Anexos: ["Anexos metodológicos", "Anexos de experiencia", "Referencias"],
};

const PREFIJO_SECCION = "mx-seccion-";
const PREFIJO_SUBSECCION = "mx-subseccion-";
const CLAVE_SECCIONES_CUSTOM = "mx-secciones-custom";

function obtenerSeccionesCustom() {
  try {
    const guardadas = localStorage.getItem(CLAVE_SECCIONES_CUSTOM);
    return guardadas ? JSON.parse(guardadas) : {};
  } catch (err) {
    return {};
  }
}

function guardarSeccionesCustom(obj) {
  try {
    localStorage.setItem(CLAVE_SECCIONES_CUSTOM, JSON.stringify(obj));
  } catch (err) {
    console.error("No se pudo guardar la sección nueva:", err);
  }
}

function buscarClaveExistente(obj, nombre) {
  const limpio = nombre.trim().toLowerCase();
  return Object.keys(obj).find((k) => k.toLowerCase() === limpio) || null;
}

/** Devuelve { "Sección": ["Sub1", "Sub2", ...], ... } combinando las
 * secciones por defecto con las que el usuario haya agregado a mano. */
function seccionesDisponibles() {
  const custom = obtenerSeccionesCustom();
  const clavesDefault = Object.keys(SECCIONES_DEFAULT);
  const clavesCustomNuevas = Object.keys(custom).filter(
    (k) => !buscarClaveExistente(SECCIONES_DEFAULT, k)
  );
  const todasLasClaves = clavesDefault.concat(clavesCustomNuevas);

  const resultado = {};
  todasLasClaves.forEach((clave) => {
    const base = SECCIONES_DEFAULT[clave] || [];
    const claveCustom = buscarClaveExistente(custom, clave);
    const extra = claveCustom ? custom[claveCustom] : [];
    const combinadas = base.concat(
      extra.filter((s) => !base.some((b) => b.toLowerCase() === s.toLowerCase()))
    );
    resultado[clave] = combinadas;
  });
  return resultado;
}

function agregarSeccionCustom(nombre) {
  const limpio = nombre.trim();
  if (!limpio) return null;

  const disponibles = seccionesDisponibles();
  const existente = buscarClaveExistente(disponibles, limpio);
  if (existente) return existente;

  const custom = obtenerSeccionesCustom();
  custom[limpio] = custom[limpio] || [];
  guardarSeccionesCustom(custom);
  return limpio;
}

function agregarSubseccionCustom(seccion, nombre) {
  const limpio = nombre.trim();
  if (!limpio || !seccion) return null;

  const disponibles = seccionesDisponibles();
  const subActuales = disponibles[seccion] || [];
  const existente = subActuales.find((s) => s.toLowerCase() === limpio.toLowerCase());
  if (existente) return existente;

  const custom = obtenerSeccionesCustom();
  const claveReal = buscarClaveExistente(custom, seccion) || seccion;
  custom[claveReal] = (custom[claveReal] || []).concat([limpio]);
  guardarSeccionesCustom(custom);
  return limpio;
}

function obtenerSeccionManual(slideId) {
  try {
    return localStorage.getItem(PREFIJO_SECCION + slideId) || null;
  } catch (err) {
    return null;
  }
}

function guardarSeccionManual(slideId, seccion) {
  try {
    localStorage.setItem(PREFIJO_SECCION + slideId, seccion);
    // La subsección anterior queda huérfana si cambia la sección --
    // se limpia para no dejar una combinación inconsistente.
    localStorage.removeItem(PREFIJO_SUBSECCION + slideId);
  } catch (err) {
    console.error("No se pudo guardar en localStorage:", err);
  }
}

function quitarSeccionManual(slideId) {
  try {
    localStorage.removeItem(PREFIJO_SECCION + slideId);
    localStorage.removeItem(PREFIJO_SUBSECCION + slideId);
  } catch (err) {
    console.error(err);
  }
}

function obtenerSubseccionManual(slideId) {
  try {
    return localStorage.getItem(PREFIJO_SUBSECCION + slideId) || null;
  } catch (err) {
    return null;
  }
}

function guardarSubseccionManual(slideId, subseccion) {
  try {
    localStorage.setItem(PREFIJO_SUBSECCION + slideId, subseccion);
  } catch (err) {
    console.error("No se pudo guardar en localStorage:", err);
  }
}

function quitarSubseccionManual(slideId) {
  try {
    localStorage.removeItem(PREFIJO_SUBSECCION + slideId);
  } catch (err) {
    console.error(err);
  }
}

function seccionEfectiva(lamina) {
  return obtenerSeccionManual(lamina.slide_id);
}

function subseccionEfectiva(lamina) {
  return obtenerSubseccionManual(lamina.slide_id);
}

/* ---------- N° de elementos (manual, opcional) ----------
   Atributo aparte de Sección/Subsección: cuántos conceptos, columnas,
   niveles, etc. tiene la lámina (ej. "Pirámide de 4 niveles"). Se asigna
   a mano igual que el nombre personalizado, con chips en el modal. */

const PREFIJO_ELEMENTOS = "mx-elementos-";

function obtenerElementosManual(slideId) {
  try {
    return localStorage.getItem(PREFIJO_ELEMENTOS + slideId) || null;
  } catch (err) {
    return null;
  }
}

function guardarElementosManual(slideId, valor) {
  try {
    localStorage.setItem(PREFIJO_ELEMENTOS + slideId, valor);
  } catch (err) {
    console.error("No se pudo guardar el N° de elementos:", err);
  }
}

function quitarElementosManual(slideId) {
  try {
    localStorage.removeItem(PREFIJO_ELEMENTOS + slideId);
  } catch (err) {
    console.error(err);
  }
}

function elementosEfectivo(lamina) {
  return obtenerElementosManual(lamina.slide_id);
}

/* ---------- Contador de uso ---------- */

const PREFIJO_USOS = "mx-usos-";

function obtenerUsos(slideId) {
  try {
    return parseInt(localStorage.getItem(PREFIJO_USOS + slideId) || "0", 10) || 0;
  } catch (err) {
    return 0;
  }
}

function incrementarUsos(slideId) {
  const total = obtenerUsos(slideId) + 1;
  try {
    localStorage.setItem(PREFIJO_USOS + slideId, String(total));
  } catch (err) {
    console.error("No se pudo guardar el contador de uso:", err);
  }
  return total;
}

/* ---------- Historial de insertadas recientemente ---------- */

const CLAVE_HISTORIAL = "mx-historial-recientes";
const HISTORIAL_MAX = 8;

function obtenerHistorial() {
  try {
    const guardado = localStorage.getItem(CLAVE_HISTORIAL);
    return guardado ? JSON.parse(guardado) : [];
  } catch (err) {
    return [];
  }
}

function registrarEnHistorial(slideId) {
  try {
    const id = String(slideId);
    let historial = obtenerHistorial().filter((x) => x !== id);
    historial.unshift(id);
    historial = historial.slice(0, HISTORIAL_MAX);
    localStorage.setItem(CLAVE_HISTORIAL, JSON.stringify(historial));
  } catch (err) {
    console.error("No se pudo guardar el historial:", err);
  }
}

function limpiarHistorial() {
  try {
    localStorage.removeItem(CLAVE_HISTORIAL);
  } catch (err) {
    console.error(err);
  }
}

/* ---------- Migración desde el sistema viejo de "Categoría" ----------
   El sistema anterior guardaba una sola categoría plana por lámina en
   localStorage bajo "mx-categoria-<slide_id>", más una lista de
   categorías personalizadas en "mx-categorias-custom". Ese trabajo no
   se pierde: se migra a Sección (sin subsección todavía) la primera vez
   que se carga el panel nuevo, para que el usuario lo reparta en
   subsecciones desde aquí mismo en vez de perderlo o repetirlo. Es
   idempotente -- no vuelve a tocar una lámina que ya tenga una Sección
   nueva asignada (sea porque ya migró antes, o porque el usuario ya la
   reclasificó a mano). */

const PREFIJO_CATEGORIA_VIEJA = "mx-categoria-";
const CLAVE_CATEGORIAS_CUSTOM_VIEJA = "mx-categorias-custom";

function migrarCategoriasViejas() {
  try {
    const customViejas = localStorage.getItem(CLAVE_CATEGORIAS_CUSTOM_VIEJA);
    if (customViejas) {
      JSON.parse(customViejas).forEach((nombre) => agregarSeccionCustom(nombre));
    }

    const clavesTotales = [];
    for (let i = 0; i < localStorage.length; i++) {
      clavesTotales.push(localStorage.key(i));
    }

    clavesTotales.forEach((clave) => {
      if (!clave || clave.indexOf(PREFIJO_CATEGORIA_VIEJA) !== 0) return;
      const slideId = clave.slice(PREFIJO_CATEGORIA_VIEJA.length);
      if (obtenerSeccionManual(slideId)) return; // ya tiene sección nueva, no tocar

      const valorViejo = localStorage.getItem(clave);
      if (!valorViejo) return;

      const seccion = agregarSeccionCustom(valorViejo);
      if (seccion) guardarSeccionManual(slideId, seccion);
    });
  } catch (err) {
    console.error("No se pudo migrar las categorías del sistema anterior:", err);
  }
}

const PREFIJO_NOMBRE = "mx-nombre-";

function obtenerNombrePersonalizado(slideId) {
  try {
    return localStorage.getItem(PREFIJO_NOMBRE + slideId) || null;
  } catch (err) {
    return null;
  }
}

function guardarNombrePersonalizado(slideId, nombre) {
  try {
    if (nombre.trim()) {
      localStorage.setItem(PREFIJO_NOMBRE + slideId, nombre.trim());
    } else {
      localStorage.removeItem(PREFIJO_NOMBRE + slideId);
    }
  } catch (err) {
    console.error("No se pudo guardar el nombre:", err);
  }
}

function nombreEfectivo(lamina) {
  return obtenerNombrePersonalizado(lamina.slide_id) || lamina.titulo_corto || lamina.titulo;
}

let indice = [];
let filtroGraficosActual = "cualquiera";
let filtroElementosActual = "cualquiera";
let filtroSeccionActual = "cualquiera";
let filtroSubseccionActual = "cualquiera";
let ordenActual = "numero";
const SIN_CATEGORIZAR = "__sin_categorizar__";
let maestroBase64 = null;
let cargandoMaestro = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    migrarCategoriasViejas();
    cargarIndice();

    // Cada grupo de chips limpia "activo" solo dentro de sí mismo, para que
    // el grupo de N° de gráficos y el de N° de elementos no se pisen entre sí.
    document.querySelectorAll("#botones-graficos .mx-chip").forEach((chip) => {
      chip.addEventListener("click", () => {
        chip.parentElement.querySelectorAll(".mx-chip").forEach((c) => c.classList.remove("activo"));
        chip.classList.add("activo");
        filtroGraficosActual = chip.getAttribute("data-valor");
        renderizar();
      });
    });

    document.querySelectorAll("#botones-elementos .mx-chip").forEach((chip) => {
      chip.addEventListener("click", () => {
        chip.parentElement.querySelectorAll(".mx-chip").forEach((c) => c.classList.remove("activo"));
        chip.classList.add("activo");
        filtroElementosActual = chip.getAttribute("data-valor");
        renderizar();
      });
    });

    document.getElementById("filtro-texto").addEventListener("input", renderizar);

    document.getElementById("orden-select").addEventListener("change", (ev) => {
      ordenActual = ev.target.value;
      renderizar();
    });

    document.getElementById("btn-limpiar-recientes").addEventListener("click", () => {
      limpiarHistorial();
      renderizarRecientes();
    });

    document.getElementById("btn-limpiar-filtros").addEventListener("click", limpiarFiltros);

    document.getElementById("modal-backdrop").addEventListener("click", (ev) => {
      if (ev.target.id === "modal-backdrop") cerrarModal();
    });

    document.getElementById("btn-exportar-secciones").addEventListener("click", exportarSecciones);

    document.getElementById("btn-importar-secciones").addEventListener("click", () => {
      document.getElementById("input-importar-secciones").click();
    });
    document.getElementById("input-importar-secciones").addEventListener("change", (ev) => {
      const archivo = ev.target.files[0];
      if (archivo) importarSecciones(archivo);
      ev.target.value = ""; // permite volver a elegir el mismo archivo después
    });
  }
});

async function cargarIndice() {
  try {
    const resp = await fetch("/indice.json");
    indice = await resp.json();
    renderizar();
  } catch (err) {
    document.getElementById("resultados").innerHTML =
      '<div class="mx-empty">No se pudo cargar el índice.<br/>' + err.message + "</div>";
  }
}

function seccionesEnUso() {
  const set = new Set();
  indice.forEach((lamina) => {
    const s = seccionEfectiva(lamina);
    if (s) set.add(s);
  });
  return Array.from(set).sort((a, b) => a.localeCompare(b, "es"));
}

function subseccionesEnUso(seccion) {
  const set = new Set();
  indice.forEach((lamina) => {
    if (seccionEfectiva(lamina) === seccion) {
      const sub = subseccionEfectiva(lamina);
      if (sub) set.add(sub);
    }
  });
  return Array.from(set).sort((a, b) => a.localeCompare(b, "es"));
}

function renderizarFiltroSeccion() {
  const select = document.getElementById("filtro-seccion-select");
  const opciones = [{ valor: "cualquiera", etiqueta: "Cualquiera" }].concat(
    seccionesEnUso().map((s) => ({ valor: s, etiqueta: s }))
  );
  opciones.push({ valor: SIN_CATEGORIZAR, etiqueta: "Sin sección" });

  select.innerHTML = opciones
    .map((op) => {
      const seleccionado = op.valor === filtroSeccionActual ? " selected" : "";
      return '<option value="' + escapeHtml(op.valor) + '"' + seleccionado + ">" + escapeHtml(op.etiqueta) + "</option>";
    })
    .join("");

  select.onchange = () => {
    filtroSeccionActual = select.value;
    filtroSubseccionActual = "cualquiera";
    renderizar();
  };
}

function renderizarFiltroSubseccion() {
  const select = document.getElementById("filtro-subseccion-select");
  const haySeccionConcreta =
    filtroSeccionActual !== "cualquiera" && filtroSeccionActual !== SIN_CATEGORIZAR;

  if (!haySeccionConcreta) {
    select.innerHTML = '<option value="cualquiera">Elige una sección primero</option>';
    select.disabled = true;
    select.onchange = null;
    return;
  }

  select.disabled = false;
  const opciones = [{ valor: "cualquiera", etiqueta: "Cualquiera (todas las de esta sección)" }].concat(
    subseccionesEnUso(filtroSeccionActual).map((s) => ({ valor: s, etiqueta: s }))
  );
  opciones.push({ valor: SIN_CATEGORIZAR, etiqueta: "Sin subsección" });

  select.innerHTML = opciones
    .map((op) => {
      const seleccionado = op.valor === filtroSubseccionActual ? " selected" : "";
      return '<option value="' + escapeHtml(op.valor) + '"' + seleccionado + ">" + escapeHtml(op.etiqueta) + "</option>";
    })
    .join("");

  select.onchange = () => {
    filtroSubseccionActual = select.value;
    renderizar();
  };
}

/** Vuelve todos los filtros a su estado inicial: gráficos, elementos,
 * sección, subsección, texto y orden — dejando la lista tal como se ve
 * al abrir el panel por primera vez (1 a 287, sin ordenar por uso). */
function limpiarFiltros() {
  filtroGraficosActual = "cualquiera";
  filtroElementosActual = "cualquiera";
  filtroSeccionActual = "cualquiera";
  filtroSubseccionActual = "cualquiera";
  ordenActual = "numero";

  document.getElementById("filtro-texto").value = "";
  document.getElementById("orden-select").value = "numero";

  ["#botones-graficos", "#botones-elementos"].forEach((sel) => {
    document.querySelectorAll(sel + " .mx-chip").forEach((c) => c.classList.remove("activo"));
    const porDefecto = document.querySelector(sel + ' .mx-chip[data-valor="cualquiera"]');
    if (porDefecto) porDefecto.classList.add("activo");
  });

  renderizar();
}

function renderizar() {
  renderizarFiltroSeccion();
  renderizarFiltroSubseccion();
  renderizarRecientes();

  const filtroTexto = document.getElementById("filtro-texto").value.toLowerCase();

  const filtradas = indice.filter((lamina) => {
    let pasaGraficos = true;
    if (filtroGraficosActual === "0") pasaGraficos = lamina.n_graficos_total === 0;
    else if (filtroGraficosActual === "1") pasaGraficos = lamina.n_graficos_total === 1;
    else if (filtroGraficosActual === "2") pasaGraficos = lamina.n_graficos_total === 2;
    else if (filtroGraficosActual === "3+") pasaGraficos = lamina.n_graficos_total >= 3;

    let pasaElementos = true;
    if (filtroElementosActual !== "cualquiera") {
      const valor = elementosEfectivo(lamina);
      if (!valor) pasaElementos = false;
      else if (filtroElementosActual === "6+") {
        const numero = parseInt(valor, 10);
        pasaElementos = valor === "6+" || (!isNaN(numero) && numero >= 6);
      } else {
        pasaElementos = valor === filtroElementosActual;
      }
    }

    let pasaSeccion = true;
    if (filtroSeccionActual === SIN_CATEGORIZAR) pasaSeccion = !seccionEfectiva(lamina);
    else if (filtroSeccionActual !== "cualquiera")
      pasaSeccion = seccionEfectiva(lamina) === filtroSeccionActual;

    let pasaSubseccion = true;
    if (
      pasaSeccion &&
      filtroSeccionActual !== "cualquiera" &&
      filtroSeccionActual !== SIN_CATEGORIZAR
    ) {
      if (filtroSubseccionActual === SIN_CATEGORIZAR) pasaSubseccion = !subseccionEfectiva(lamina);
      else if (filtroSubseccionActual !== "cualquiera")
        pasaSubseccion = subseccionEfectiva(lamina) === filtroSubseccionActual;
    }

    // La búsqueda de texto mira el título real, el nombre personalizado y la
    // Sección/Subsección asignadas -- así "pirámide" encuentra las láminas
    // aunque su título original diga otra cosa (ej. "Idea principal, 5 dimensiones").
    const titulo = (lamina.titulo || "").toLowerCase();
    const nombrePersonalizado = (nombreEfectivo(lamina) || "").toLowerCase();
    const seccionTexto = (seccionEfectiva(lamina) || "").toLowerCase();
    const subseccionTexto = (subseccionEfectiva(lamina) || "").toLowerCase();
    const pasaTexto =
      filtroTexto === "" ||
      titulo.includes(filtroTexto) ||
      nombrePersonalizado.includes(filtroTexto) ||
      seccionTexto.includes(filtroTexto) ||
      subseccionTexto.includes(filtroTexto);

    return pasaGraficos && pasaElementos && pasaSeccion && pasaSubseccion && pasaTexto;
  });

  if (ordenActual === "usos") {
    filtradas.sort((a, b) => {
      const diferencia = obtenerUsos(b.slide_id) - obtenerUsos(a.slide_id);
      return diferencia !== 0 ? diferencia : a.slide - b.slide;
    });
  }

  document.getElementById("contador").innerText =
    filtradas.length + (filtradas.length === 1 ? " lámina encontrada" : " láminas encontradas");

  actualizarProgreso();

  const contenedor = document.getElementById("resultados");
  contenedor.innerHTML = "";

  if (filtradas.length === 0) {
    contenedor.innerHTML =
      '<div class="mx-empty">Sin resultados con estos filtros.<br/>Ajusta el rango de gráficos, elementos, la sección o el texto de búsqueda.</div>';
    return;
  }

  filtradas.forEach((lamina) => {
    const div = document.createElement("div");
    div.className = "mx-card";

    const numFormateado = String(lamina.slide).padStart(2, "0");
    const nombreMostrado = nombreEfectivo(lamina);
    const tituloHtml = nombreMostrado
      ? '<div class="mx-card-title">' + escapeHtml(nombreMostrado) + "</div>"
      : '<div class="mx-card-title sin-titulo">(sin título)</div>';

    const visual = lamina.thumb
      ? '<div class="mx-thumb-wrap">' +
        '<img class="mx-thumb" src="' +
        escapeHtml(lamina.thumb) +
        '" alt="Vista previa de la lámina ' + lamina.slide + '" loading="lazy" />' +
        '<span class="mx-thumb-num">' + numFormateado + "</span>" +
        "</div>"
      : '<div class="mx-card-num">' + numFormateado + "</div>";

    const seccion = seccionEfectiva(lamina);
    const subseccion = subseccionEfectiva(lamina);
    let tagsSeccion;
    if (seccion) {
      tagsSeccion =
        '<span class="mx-tag">' + escapeHtml(seccion) + "</span>" +
        (subseccion ? '<span class="mx-tag">' + escapeHtml(subseccion) + "</span>" : "");
    } else {
      tagsSeccion = '<span class="mx-tag pendiente">Sin sección</span>';
    }

    const elementos = elementosEfectivo(lamina);
    const tagElementos = !elementos
      ? ""
      : elementos === "N/A"
      ? '<span class="mx-tag mx-tag-na">No aplica</span>'
      : '<span class="mx-tag">' + escapeHtml(elementos) + " elem.</span>";

    const usos = obtenerUsos(lamina.slide_id);
    const tagUsos = usos > 0
      ? '<span class="mx-tag usos">Usada ' + usos + (usos === 1 ? " vez" : " veces") + "</span>"
      : "";

    div.innerHTML =
      visual +
      '<div class="mx-card-body">' +
      tituloHtml +
      '<div class="mx-tags">' +
      tagsSeccion +
      tagElementos +
      tag(lamina.n_graficos_total, lamina.n_graficos_total === 1 ? "gráfico" : "gráficos") +
      tag(lamina.n_tablas, lamina.n_tablas === 1 ? "tabla" : "tablas") +
      tag(lamina.n_imagenes, lamina.n_imagenes === 1 ? "imagen" : "imágenes") +
      tagUsos +
      "</div>" +
      "</div>";

    div.addEventListener("click", () => abrirModal(lamina));
    contenedor.appendChild(div);
  });
}

/** Franja de acceso rápido a las últimas láminas insertadas por este
 * consultor (guardadas en este navegador/localStorage). Se oculta sola
 * si todavía no se ha insertado ninguna lámina. */
function renderizarRecientes() {
  const wrap = document.getElementById("recientes-wrap");
  const lista = document.getElementById("recientes-lista");
  const historial = obtenerHistorial();

  if (!historial.length || !indice.length) {
    wrap.style.display = "none";
    return;
  }

  const laminas = historial
    .map((id) => indice.find((l) => String(l.slide_id) === String(id)))
    .filter((l) => !!l);

  if (!laminas.length) {
    wrap.style.display = "none";
    return;
  }

  wrap.style.display = "";
  lista.innerHTML = laminas
    .map((lamina) => {
      const num = String(lamina.slide).padStart(2, "0");
      const thumb = lamina.thumb
        ? '<img class="mx-recientes-thumb" src="' + escapeHtml(lamina.thumb) +
          '" alt="Lámina ' + lamina.slide + '" loading="lazy" />'
        : '<div class="mx-recientes-thumb mx-recientes-thumb-vacio">' + num + "</div>";
      return (
        '<div class="mx-recientes-item" data-slide-id="' + escapeHtml(String(lamina.slide_id)) + '">' +
        thumb +
        "</div>"
      );
    })
    .join("");

  lista.querySelectorAll(".mx-recientes-item").forEach((item) => {
    item.addEventListener("click", () => {
      const id = item.getAttribute("data-slide-id");
      const lamina = indice.find((l) => String(l.slide_id) === id);
      if (lamina) abrirModal(lamina);
    });
  });
}

function tag(valor, etiqueta) {
  const clase = valor === 0 ? "mx-tag cero" : "mx-tag";
  return '<span class="' + clase + '">' + valor + " " + etiqueta + "</span>";
}

function escapeHtml(texto) {
  const div = document.createElement("div");
  div.innerText = texto;
  return div.innerHTML;
}

/* ---------- Modal de previsualización ---------- */

function abrirModal(lamina) {
  const contenido = document.getElementById("modal-contenido");

  const imagenHtml = lamina.thumb
    ? '<div class="mx-modal-img-wrap"><img class="mx-modal-img" src="' +
      escapeHtml(lamina.thumb) +
      '" alt="Vista previa ampliada de la lámina ' + lamina.slide + '" /></div>'
    : "";

  const tituloCompleto = lamina.titulo || "(sin título)";
  const nombreActual = nombreEfectivo(lamina);
  const seccionActiva = seccionEfectiva(lamina);
  const subseccionActiva = subseccionEfectiva(lamina);
  const elementosActivo = elementosEfectivo(lamina);
  const tieneSeccion = !!seccionActiva;

  const elementosHtml =
    '<div class="mx-cat-label">N° de elementos <span class="mx-label-hint">(opcional — conceptos, columnas, niveles… sirve para filtrar, ej. "pirámide de 4")</span></div>' +
    '<div class="mx-cat-picker" id="elementos-picker">' +
    ["1", "2", "3", "4", "5", "6", "7", "8+"]
      .map((v) => {
        const activo = v === elementosActivo ? " activo" : "";
        return '<span class="mx-cat-chip' + activo + '" data-valor="' + v + '">' + v + "</span>";
      })
      .join("") +
    '<span class="mx-cat-chip mx-cat-chip-na' +
    (elementosActivo === "N/A" ? " activo" : "") +
    '" data-valor="N/A" title="Los consultores revisaron esta lámina y el concepto de \'N° de elementos\' no aplica (ej. portada, texto narrativo, cierre)">No aplica</span>' +
    '<span class="mx-cat-chip' + (elementosActivo ? "" : " activo") + '" data-valor="">— Sin dato —</span>' +
    "</div>";

  let hintPlantilla = "";
  if (lamina.tipo_plantilla) {
    hintPlantilla =
      '<div class="mx-cat-detectada">Plantilla detectada automáticamente: ' +
      escapeHtml(lamina.tipo_plantilla) +
      " (dato informativo, no afecta la sección)</div>";
  }

  const disponibles = seccionesDisponibles();
  const clavesSeccion = Object.keys(disponibles);

  const seccionHtml =
    '<div class="mx-cat-label">Sección</div>' +
    '<select class="mx-select mx-modal-select" id="seccion-select">' +
    '<option value=""' + (tieneSeccion ? "" : " selected") + ">— Elegir sección —</option>" +
    clavesSeccion
      .map((s) => {
        const seleccionado = s === seccionActiva ? " selected" : "";
        return '<option value="' + escapeHtml(s) + '"' + seleccionado + ">" + escapeHtml(s) + "</option>";
      })
      .join("") +
    "</select>" +
    '<div class="mx-cat-nueva">' +
    '<input type="text" id="seccion-nueva-input" placeholder="Crear sección nueva…" maxlength="50" />' +
    '<button type="button" id="seccion-nueva-btn">Agregar</button>' +
    "</div>";

  let subseccionHtml;
  if (tieneSeccion) {
    const subOpciones = disponibles[seccionActiva] || [];
    subseccionHtml =
      '<div class="mx-cat-label">Subsección <span class="mx-label-hint">(opcional)</span></div>' +
      '<select class="mx-select mx-modal-select" id="subseccion-select">' +
      '<option value=""' + (subseccionActiva ? "" : " selected") + ">— Sin subsección —</option>" +
      subOpciones
        .map((s) => {
          const seleccionado = s === subseccionActiva ? " selected" : "";
          return '<option value="' + escapeHtml(s) + '"' + seleccionado + ">" + escapeHtml(s) + "</option>";
        })
        .join("") +
      "</select>" +
      '<div class="mx-cat-nueva">' +
      '<input type="text" id="subseccion-nueva-input" placeholder="Crear subsección nueva…" maxlength="50" />' +
      '<button type="button" id="subseccion-nueva-btn">Agregar</button>' +
      "</div>";
  } else {
    subseccionHtml =
      '<div class="mx-cat-label">Subsección</div>' +
      '<div class="mx-cat-detectada">Elige primero una sección arriba</div>';
  }

  const nombreHtml =
    '<div class="mx-cat-label">Nombre en el panel</div>' +
    '<div class="mx-cat-nueva">' +
    '<input type="text" id="nombre-input" value="' + escapeHtml(nombreActual || "") + '" maxlength="80" />' +
    '<button type="button" id="nombre-guardar-btn">Guardar</button>' +
    "</div>";

  contenido.innerHTML =
    '<button type="button" id="modal-cerrar-x" class="mx-modal-close" aria-label="Cerrar">×</button>' +
    imagenHtml +
    '<div class="mx-modal-body">' +
    '<div class="mx-modal-titulo">Lámina ' + lamina.slide + " · " + escapeHtml(tituloCompleto) + "</div>" +
    nombreHtml +
    hintPlantilla +
    seccionHtml +
    subseccionHtml +
    elementosHtml +
    '<div class="mx-modal-tags">' +
    tag(lamina.n_graficos_total, lamina.n_graficos_total === 1 ? "gráfico" : "gráficos") +
    tag(lamina.n_tablas, lamina.n_tablas === 1 ? "tabla" : "tablas") +
    tag(lamina.n_imagenes, lamina.n_imagenes === 1 ? "imagen" : "imágenes") +
    (obtenerUsos(lamina.slide_id) > 0
      ? '<span class="mx-tag usos">Usada ' + obtenerUsos(lamina.slide_id) +
        (obtenerUsos(lamina.slide_id) === 1 ? " vez" : " veces") + "</span>"
      : "") +
    "</div>" +
    '<div class="mx-modal-actions">' +
    '<button class="mx-btn" id="modal-cerrar">Cerrar</button>' +
    '<button class="mx-btn mx-btn-primario" id="modal-insertar">Insertar esta lámina</button>' +
    "</div>" +
    '<div id="modal-estado" style="font-size:11px; margin-top:8px; color: var(--mx-navy); opacity:0.7;"></div>' +
    "</div>";

  document.getElementById("modal-cerrar").addEventListener("click", cerrarModal);
  document.getElementById("modal-cerrar-x").addEventListener("click", cerrarModal);
  document.getElementById("modal-insertar").addEventListener("click", () => insertarLamina(lamina));

  // --- N° de elementos ---
  document.querySelectorAll("#elementos-picker .mx-cat-chip").forEach((chip) => {
    chip.addEventListener("click", () => {
      const valor = chip.getAttribute("data-valor");
      if (valor) guardarElementosManual(lamina.slide_id, valor);
      else quitarElementosManual(lamina.slide_id);
      renderizar();
      abrirModal(lamina);
    });
  });

  // --- Sección ---
  const seccionSelect = document.getElementById("seccion-select");
  seccionSelect.addEventListener("change", () => {
    const valor = seccionSelect.value;
    if (!valor) return; // "— Elegir sección —": no hacer nada
    guardarSeccionManual(lamina.slide_id, valor);
    renderizar();
    abrirModal(lamina);
  });

  const inputSeccionNueva = document.getElementById("seccion-nueva-input");
  const confirmarSeccionNueva = () => {
    const seccion = agregarSeccionCustom(inputSeccionNueva.value);
    if (!seccion) return;
    guardarSeccionManual(lamina.slide_id, seccion);
    renderizar();
    abrirModal(lamina);
  };
  document.getElementById("seccion-nueva-btn").addEventListener("click", confirmarSeccionNueva);
  inputSeccionNueva.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") {
      ev.preventDefault();
      confirmarSeccionNueva();
    }
  });

  // --- Subsección (solo si ya hay sección elegida) ---
  if (tieneSeccion) {
    const subseccionSelect = document.getElementById("subseccion-select");
    subseccionSelect.addEventListener("change", () => {
      const valor = subseccionSelect.value;
      if (valor) guardarSubseccionManual(lamina.slide_id, valor);
      else quitarSubseccionManual(lamina.slide_id);
      renderizar();
      abrirModal(lamina);
    });

    const inputSubNueva = document.getElementById("subseccion-nueva-input");
    const confirmarSubNueva = () => {
      const sub = agregarSubseccionCustom(seccionActiva, inputSubNueva.value);
      if (!sub) return;
      guardarSubseccionManual(lamina.slide_id, sub);
      renderizar();
      abrirModal(lamina);
    };
    document.getElementById("subseccion-nueva-btn").addEventListener("click", confirmarSubNueva);
    inputSubNueva.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") {
        ev.preventDefault();
        confirmarSubNueva();
      }
    });
  }

  // --- Nombre ---
  const inputNombre = document.getElementById("nombre-input");
  document.getElementById("nombre-guardar-btn").addEventListener("click", () => {
    guardarNombrePersonalizado(lamina.slide_id, inputNombre.value);
    renderizar();
    abrirModal(lamina);
  });
  inputNombre.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") {
      ev.preventDefault();
      guardarNombrePersonalizado(lamina.slide_id, inputNombre.value);
      renderizar();
      abrirModal(lamina);
    }
  });

  document.getElementById("modal-backdrop").classList.add("abierto");
}

function cerrarModal() {
  document.getElementById("modal-backdrop").classList.remove("abierto");
  document.getElementById("modal-contenido").innerHTML = "";
}

/* ---------- Progreso y exportación de secciones manuales ---------- */

function actualizarProgreso() {
  const total = indice.length;
  const conSeccion = indice.filter((l) => seccionEfectiva(l)).length;
  document.getElementById("progreso-categorias").innerText =
    conSeccion + "/" + total + " láminas con sección asignada";
}

function exportarSecciones() {
  const salida = {};
  let total = 0;
  indice.forEach((lamina) => {
    const seccion = obtenerSeccionManual(lamina.slide_id);
    if (seccion) {
      salida[String(lamina.slide_id)] = {
        seccion: seccion,
        subseccion: obtenerSubseccionManual(lamina.slide_id) || null,
        elementos: obtenerElementosManual(lamina.slide_id) || null,
      };
      total += 1;
    }
  });

  if (total === 0) {
    alert_reemplazo("Todavía no has asignado sección a ninguna lámina.");
    return;
  }

  const blob = new Blob([JSON.stringify(salida, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const enlace = document.createElement("a");
  enlace.href = url;
  enlace.download = "secciones.json";
  document.body.appendChild(enlace);
  enlace.click();
  document.body.removeChild(enlace);
  URL.revokeObjectURL(url);
}

function alert_reemplazo(mensaje) {
  // Nunca usar alert() nativo dentro de un taskpane de Office (bloquea el
  // WebView y lanza "Uncaught runtime errors"). Se muestra en el DOM.
  document.getElementById("progreso-categorias").innerText = mensaje;
}

/** Importa un secciones.json (mismo formato que exportarSecciones genera:
 * { "<slide_id>": { "seccion": "...", "subseccion": "..." | null }, ... } )
 * y lo aplica a localStorage. Sobrescribe la sección/subsección de cualquier
 * lámina que ya estuviera clasificada. Cualquier sección o subsección que
 * no exista todavía en este panel se agrega como personalizada, para que
 * aparezca en los desplegables igual que si se hubiera creado a mano. */
function importarSecciones(archivo) {
  const lector = new FileReader();
  lector.onload = () => {
    let datos;
    try {
      datos = JSON.parse(lector.result);
    } catch (err) {
      alert_reemplazo("El archivo elegido no es un JSON válido.");
      return;
    }

    let aplicadas = 0;
    Object.keys(datos).forEach((slideId) => {
      const entrada = datos[slideId];
      if (!entrada || !entrada.seccion) return;

      const seccion = agregarSeccionCustom(entrada.seccion);
      if (!seccion) return;
      guardarSeccionManual(slideId, seccion); // esto también limpia la subsección vieja

      if (entrada.subseccion) {
        const subseccion = agregarSubseccionCustom(seccion, entrada.subseccion);
        if (subseccion) guardarSubseccionManual(slideId, subseccion);
      }
      if (entrada.elementos) {
        guardarElementosManual(slideId, String(entrada.elementos));
      }
      aplicadas += 1;
    });

    renderizar();
    alert_reemplazo(
      aplicadas > 0
        ? "Importadas " + aplicadas + " láminas desde el archivo."
        : "El archivo no tenía láminas con sección asignada."
    );
  };
  lector.onerror = () => alert_reemplazo("No se pudo leer el archivo elegido.");
  lector.readAsText(archivo);
}

async function obtenerMaestroBase64() {
  if (maestroBase64) return maestroBase64;
  if (cargandoMaestro) return cargandoMaestro;

  cargandoMaestro = (async () => {
    const resp = await fetch(RUTA_MAESTRO);
    if (!resp.ok) {
      throw new Error("No se encontró " + RUTA_MAESTRO + " (HTTP " + resp.status + ")");
    }
    const blob = await resp.blob();

    const base64 = await new Promise((resolve, reject) => {
      const lector = new FileReader();
      lector.onload = () => {
        const resultado = lector.result;
        const coma = resultado.indexOf(",");
        resolve(coma >= 0 ? resultado.slice(coma + 1) : resultado);
      };
      lector.onerror = () => reject(new Error("Falló la lectura del archivo maestro"));
      lector.readAsDataURL(blob);
    });

    maestroBase64 = base64;
    cargandoMaestro = null;
    return base64;
  })();

  return cargandoMaestro;
}

async function insertarLamina(lamina) {
  const botonInsertar = document.getElementById("modal-insertar");
  const botonCerrar = document.getElementById("modal-cerrar");
  const estado = document.getElementById("modal-estado");

  botonInsertar.disabled = true;
  botonCerrar.disabled = true;
  estado.innerText = maestroBase64
    ? "Insertando…"
    : "Preparando el archivo maestro por primera vez. Puede tardar un momento…";

  try {
    const base64 = await obtenerMaestroBase64();

    await PowerPoint.run(async (context) => {
      const presentacion = context.presentation;

      const seleccionada = presentacion.getSelectedSlides().getItemAt(0);
      seleccionada.load("id");
      await context.sync();

      const opciones = {
        formatting: "KeepSourceFormatting",
        sourceSlideIds: [String(lamina.slide_id) + "#"],
        targetSlideId: seleccionada.id,
      };

      presentacion.insertSlidesFromBase64(base64, opciones);
      await context.sync();
    });

    estado.innerText = "Insertada correctamente.";
    incrementarUsos(lamina.slide_id);
    registrarEnHistorial(lamina.slide_id);
    renderizar();
    setTimeout(cerrarModal, 900);
  } catch (err) {
    let detalle = err.message || String(err);
    if (err.code === "SlideNotFound" || /SlideNotFound/i.test(detalle)) {
      detalle =
        "PowerPoint no encontró la lámina en el archivo maestro (slide_id=" +
        lamina.slide_id +
        "). Revisa que el indice.json corresponda al maestro que está en public/.";
    }
    estado.innerText = "No se pudo insertar. " + detalle;
    botonInsertar.disabled = false;
    botonCerrar.disabled = false;
  }
}