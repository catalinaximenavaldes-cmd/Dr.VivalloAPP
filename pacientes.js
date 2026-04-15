const URL =
  typeof CONFIG !== "undefined" && CONFIG.API_URL
    ? CONFIG.API_URL
    : "https://script.google.com/macros/s/AKfycbz_o5PPXlSY8Q7kcIwKjVETz5m-lhU7TxZO84bo9OSE8zf8qsB0Fd6kijopc4gWy94/exec";

let pacientes = {};
let historialPacientes = {};
let diagnosticosMap = {};
let pacienteEditando = "";
let historialCargado = false;
let historialCargaPromise = null;

function escapeHtml(texto) {
  return String(texto ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function normalizarFechaInput(valor) {
  if (!valor) return "";
  if (typeof valor === "string") {
    if (/^\d{4}-\d{2}-\d{2}$/.test(valor)) return valor;
    if (/^\d{4}-\d{2}-\d{2}T/.test(valor)) return valor.slice(0, 10);
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(valor)) {
      const [dia, mes, anio] = valor.split("/");
      return `${anio}-${mes.padStart(2, "0")}-${dia.padStart(2, "0")}`;
    }
  }
  const fecha = new Date(valor);
  if (Number.isNaN(fecha.getTime())) return "";
  const anio = fecha.getFullYear();
  const mes = String(fecha.getMonth() + 1).padStart(2, "0");
  const dia = String(fecha.getDate()).padStart(2, "0");
  return `${anio}-${mes}-${dia}`;
}

function normalizarFechaRegistro(valor) {
  const fecha = new Date(valor);
  return Number.isNaN(fecha.getTime()) ? new Date() : fecha;
}

function formatearFechaHora(valor) {
  const fecha = valor instanceof Date ? valor : new Date(valor);
  if (Number.isNaN(fecha.getTime())) return "Sin fecha";
  return fecha.toLocaleString("es-CL", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function formatearFechaNacimiento(valor) {
  const fechaIso = normalizarFechaInput(valor);
  if (!fechaIso) return "No registrada";
  const [anio, mes, dia] = fechaIso.split("-");
  return `${dia}/${mes}/${anio}`;
}

function calcularEdadDesdeFecha(valor) {
  const fechaIso = normalizarFechaInput(valor);
  if (!fechaIso) return "";
  const [anio, mes, dia] = fechaIso.split("-").map(Number);
  const hoy = new Date();
  let edad = hoy.getFullYear() - anio;
  const mesActual = hoy.getMonth() + 1;
  const diaActual = hoy.getDate();
  if (mesActual < mes || (mesActual === mes && diaActual < dia)) edad--;
  return edad >= 0 ? edad : "";
}

function agregarAlHistorial(nombre, item) {
  if (!nombre) return;
  if (!historialPacientes[nombre]) historialPacientes[nombre] = [];
  historialPacientes[nombre].push(item);
}

function reconstruirHistorialDesdePacientes() {
  historialPacientes = {};

  Object.keys(pacientes).forEach(nombre => {
    const registros = [...(pacientes[nombre] || [])].sort((a, b) => a.fecha - b.fecha);
    registros.forEach((registro, index) => {
      if (!registro.evolucion) return;
      agregarAlHistorial(nombre, {
        fecha: registro.fecha,
        evolucion: registro.evolucion,
        tabla: "Hoja 1",
        tipo: index === 0 ? "Ingreso" : "Evolucion"
      });
    });
  });
}

function getPrimerRegistro(nombre) {
  const lista = pacientes[nombre] || [];
  if (!lista.length) return null;
  return [...lista].sort((a, b) => a.fecha - b.fecha)[0];
}

function getDiagnosticoActual(nombre) {
  return diagnosticosMap[nombre] || "";
}

async function cargarBaseDatos() {
  const lista = document.getElementById("lista");

  try {
    const resPacientes = await fetch(URL);
    const dataPacientes = await resPacientes.json();

    if (!Array.isArray(dataPacientes)) {
      throw new Error("Respuesta invalida del servidor");
    }

    pacientes = {};
    historialPacientes = {};
    diagnosticosMap = {};
    historialCargado = false;
    historialCargaPromise = null;

    for (let i = 1; i < dataPacientes.length; i++) {
      const fila = dataPacientes[i] || [];
      const nombre = (fila[1] || "").toString().trim();
      if (!nombre) continue;

      if (!pacientes[nombre]) pacientes[nombre] = [];

      pacientes[nombre].push({
        fecha: normalizarFechaRegistro(fila[0]),
        rut: fila[2] || "",
        fechaNacimiento: fila[3] || "",
        edad: fila[4] || "",
        direccion: fila[5] || "",
        telefono: fila[6] || "",
        correo: fila[7] || "",
        evolucion: fila[8] || "",
        tratamientoActivo: fila[9] || ""
      });
    }

    reconstruirHistorialDesdePacientes();
    mostrarLista();

    const pacienteGuardado = localStorage.getItem("pacienteActual") || "";
    const volverAFichaPaciente = localStorage.getItem("volverAFichaPaciente") === "1";
    if (volverAFichaPaciente && pacienteGuardado && pacientes[pacienteGuardado]) {
      localStorage.removeItem("volverAFichaPaciente");
      verPaciente(pacienteGuardado);
    }
  } catch (error) {
    console.error("Error al conectar:", error);
    if (lista) lista.innerHTML = "Error al conectar.";
  }
}

async function cargarHistorialCompleto() {
  if (historialCargado) return;
  if (historialCargaPromise) {
    await historialCargaPromise;
    return;
  }

  historialCargaPromise = (async () => {
    reconstruirHistorialDesdePacientes();

    const pedidos = [
      fetch(URL + "?tabla=Documentos").then(r => r.json()).catch(() => []),
      fetch(URL + "?tabla=Diagnosticos").then(r => r.json()).catch(() => [])
    ];

    const [dataDocumentos, dataDiagnosticos] = await Promise.all(pedidos);

    if (Array.isArray(dataDocumentos)) {
      for (let i = 1; i < dataDocumentos.length; i++) {
        const fila = dataDocumentos[i] || [];
        agregarAlHistorial(fila[1], {
          fecha: normalizarFechaRegistro(fila[0]),
          evolucion: fila[4] || "",
          tabla: "Documentos",
          tipo: fila[3] || "Documento"
        });
      }
    }

    if (Array.isArray(dataDiagnosticos)) {
      diagnosticosMap = {};
      for (let i = 1; i < dataDiagnosticos.length; i++) {
        const fila = dataDiagnosticos[i] || [];
        const nombre = (fila[0] || "").toString().trim();
        if (!nombre) continue;
        diagnosticosMap[nombre] = fila[1] || "";
      }
    }

    historialCargado = true;
  })();

  try {
    await historialCargaPromise;
  } finally {
    historialCargaPromise = null;
  }
}

function mostrarLista() {
  const lista = document.getElementById("lista");
  document.getElementById("btnVolverLista").style.display = "none";
  document.getElementById("titulo-pagina").innerText = "Pacientes ingresados";

  const nombres = Object.keys(pacientes).sort((a, b) => a.localeCompare(b, "es"));
  if (!nombres.length) {
    lista.innerHTML = "<p>No hay pacientes.</p>";
    return;
  }

  lista.innerHTML = nombres
    .map(nombre => `<div class="card-paciente" onclick="verPaciente('${escapeHtml(nombre)}')">${escapeHtml(nombre)}</div>`)
    .join("");
}

async function verPaciente(nombre) {
  const perfil = getPrimerRegistro(nombre);
  if (!perfil) return;

  localStorage.setItem("pacienteActual", nombre);
  document.getElementById("btnVolverLista").style.display = "block";
  document.getElementById("titulo-pagina").innerText = "Expediente Completo";

  document.getElementById("lista").innerHTML = `
    <div class="perfil-paciente">
      <h3 style="display:flex; justify-content:space-between; align-items:center;">
        <span>Ficha: ${escapeHtml(nombre)}</span>
        <button class="btn-editar-ficha" onclick="abrirModalEdicion('${escapeHtml(nombre)}')" title="Editar ficha" aria-label="Editar ficha">&#9998;</button>
      </h3>
      <div class="datos-personales">
        <p><strong>RUT:</strong> ${escapeHtml(perfil.rut || "No registrado")}</p>
        <p><strong>Fecha de nacimiento:</strong> ${escapeHtml(formatearFechaNacimiento(perfil.fechaNacimiento))}</p>
        <p><strong>Edad:</strong> ${escapeHtml(String(perfil.edad || calcularEdadDesdeFecha(perfil.fechaNacimiento) || "No registrada"))} anos</p>
        <p><strong>Direccion:</strong> ${escapeHtml(perfil.direccion || "No registrada")}</p>
        <p><strong>Telefono:</strong> ${escapeHtml(perfil.telefono || "No registrado")}</p>
        <p><strong>Correo:</strong> ${escapeHtml(perfil.correo || "No registrado")}</p>
        <div class="contenedor-diagnostico no-print">
          <h4>Diagnostico Principal</h4>
          <textarea id="txt-diagnostico" placeholder="Cargando..."></textarea>
          <button onclick="guardarDiagnostico('${escapeHtml(nombre)}')" style="background:#2980b9; color:white; padding:10px; border-radius:8px; margin-top:5px; border:none; width:100%; font-weight:bold; cursor:pointer;">Guardar Diagnostico</button>
        </div>
      </div>
      <div class="columna-botones no-print">
        <button class="btn-guardar" onclick="window.location.href='evolución.html'">Nueva Evolucion</button>
        <button class="btn-documentos" onclick="window.location.href='documentos.html'">Recetas y Examenes</button>
        <button onclick="window.location.href='imagenes-paciente.html'" style="background:#8e44ad; color:white;">Imagenes</button>
        <button onclick="altaPaciente('${escapeHtml(nombre)}')" style="background:#e67e22; color:white;">Dar de Alta</button>
      </div>
    </div>
    <div id="historial-atenciones">
      <h4 style="color:#2c3e50; margin: 20px 0 10px 10px;" class="no-print">Historial de Atenciones</h4>
      <p style="margin:10px;" class="no-print">Cargando historial...</p>
    </div>
  `;

  abrirModalEdicionEstado(nombre);
  cargarDiagnostico(nombre);

  try {
    await cargarHistorialCompleto();
  } catch (error) {
    console.error("Error al cargar historial:", error);
  }

  renderizarHistorial(nombre);
}

function renderizarHistorial(nombre) {
  const contenedor = document.getElementById("historial-atenciones");
  if (!contenedor) return;

  const registrosCronologicos = [...(historialPacientes[nombre] || [])].sort((a, b) => a.fecha - b.fecha);
  const registros = [...registrosCronologicos].reverse();
  let html = '<h4 style="color:#2c3e50; margin: 20px 0 10px 10px;" class="no-print">Historial de Atenciones</h4>';

  if (!registros.length) {
    contenedor.innerHTML = html + '<p style="margin:10px;">No hay evoluciones ni documentos registrados.</p>';
    return;
  }

  html += registros
    .map((registro, index) => {
      const abierta = index === 0;
      const numero = registrosCronologicos.findIndex(item => item === registro) + 1;
      return `
        <div class="contenedor-acordeon">
          <div class="header-nota ${abierta ? "abierto" : ""}" onclick="toggleNota(this)">
            <span>${numero}- ${escapeHtml(registro.tipo || "Registro")} - ${escapeHtml(formatearFechaHora(registro.fecha))}</span>
            <span>&#9662;</span>
          </div>
          <div class="contenido-nota" style="display:${abierta ? "block" : "none"};">
            <p style="white-space:pre-wrap;">${escapeHtml(registro.evolucion || "")}</p>
            <button class="no-print" onclick="eliminarNota('${escapeHtml(nombre)}', '${registro.fecha.toISOString()}', '${escapeHtml(registro.tabla || "Evoluciones")}')" style="background:none; border:none; color:red; cursor:pointer; font-size:12px; padding:0; text-decoration:underline;">Eliminar entrada</button>
          </div>
        </div>
      `;
    })
    .join("");

  contenedor.innerHTML = html;
}

function abrirModalEdicionEstado(nombre) {
  const perfil = getPrimerRegistro(nombre);
  if (!perfil) return;
  pacienteEditando = perfil.rut || "";
}

function abrirModalEdicion(nombre) {
  const perfil = getPrimerRegistro(nombre);
  if (!perfil) return;

  pacienteEditando = perfil.rut || "";
  document.getElementById("modal-nombre").value = nombre;
  document.getElementById("modal-rut").value = perfil.rut || "";
  document.getElementById("modal-fecha-nacimiento").value = normalizarFechaInput(perfil.fechaNacimiento);
  document.getElementById("modal-direccion").value = perfil.direccion || "";
  document.getElementById("modal-telefono").value = perfil.telefono || "";
  document.getElementById("modal-correo").value = perfil.correo || "";
  document.getElementById("modal-edicion").style.display = "block";
}

function cerrarModal() {
  document.getElementById("modal-edicion").style.display = "none";
}

async function guardarEdicionPaciente() {
  const datos = {
    accion: "editarFichaCompleta",
    rutOriginal: pacienteEditando,
    nombre: document.getElementById("modal-nombre").value.trim(),
    rut: document.getElementById("modal-rut").value.trim(),
    fechaNac: document.getElementById("modal-fecha-nacimiento").value,
    direccion: document.getElementById("modal-direccion").value.trim(),
    telefono: document.getElementById("modal-telefono").value.trim(),
    correo: document.getElementById("modal-correo").value.trim()
  };

  try {
    await fetch(URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(datos)
    });

    alert("Paciente actualizado");
    cerrarModal();
    location.reload();
  } catch (error) {
    alert("No se pudo actualizar el paciente");
  }
}

async function cargarDiagnostico(nombre) {
  const textarea = document.getElementById("txt-diagnostico");
  if (!textarea) return;

  textarea.value = getDiagnosticoActual(nombre);
  textarea.placeholder = "Escriba el diagnostico principal...";

  if (historialCargado) return;

  try {
    await cargarHistorialCompleto();
    textarea.value = getDiagnosticoActual(nombre);
  } catch (error) {
    console.error("Error al cargar diagnostico:", error);
    textarea.placeholder = "No se pudo cargar el diagnostico";
  }
}

async function guardarDiagnostico(nombre) {
  const textarea = document.getElementById("txt-diagnostico");
  const diagnostico = textarea ? textarea.value.trim() : "";

  try {
    await fetch(URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        accion: "actualizarDiagnostico",
        nombre,
        diagnostico
      })
    });

    diagnosticosMap[nombre] = diagnostico;
    alert("Diagnostico actualizado");
  } catch (error) {
    alert("No se pudo guardar el diagnostico");
  }
}

function toggleNota(elemento) {
  elemento.classList.toggle("abierto");
  const contenido = elemento.nextElementSibling;
  if (contenido) {
    contenido.style.display = contenido.style.display === "block" ? "none" : "block";
  }
}

async function eliminarNota(nombre, fecha, tabla = "Evoluciones") {
  if (!confirm("Deseas eliminar esta entrada del historial?")) {
    return;
  }

  try {
    await fetch(URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        accion: "eliminarNota",
        nombre,
        fecha,
        tabla
      })
    });

    alert("Entrada eliminada.");
    location.reload();
  } catch (error) {
    alert("No se pudo eliminar la entrada.");
  }
}

async function altaPaciente(nombre) {
  const perfil = getPrimerRegistro(nombre);
  if (!perfil) {
    alert("No se encontro el paciente.");
    return;
  }

  if (!confirm(`Dar de alta a ${nombre} eliminara su ficha para siempre. Esta accion no se puede deshacer.`)) {
    return;
  }

  try {
    await fetch(URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        accion: "alta",
        nombre,
        rut: perfil.rut || ""
      })
    });

    alert("Paciente eliminado definitivamente.");
    location.reload();
  } catch (error) {
    alert("No se pudo eliminar al paciente.");
  }
}

async function eliminarPacienteCompleto(nombre) {
  const perfil = getPrimerRegistro(nombre);
  if (!perfil) {
    alert("No se encontro el paciente.");
    return;
  }

  if (!confirm(`Deseas eliminar por completo a ${nombre}? Esta accion borrara toda su ficha.`)) {
    return;
  }

  try {
    await fetch(URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        accion: "eliminarCompleto",
        nombre,
        rut: perfil.rut || ""
      })
    });

    alert("Paciente eliminado completamente.");
    location.reload();
  } catch (error) {
    alert("No se pudo eliminar el paciente.");
  }
}

cargarBaseDatos();

