// ==================== RECIBIR PETICIONES ====================
function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (error) {
    data = e.parameter;
  }

  let respuesta = {};

  switch (data.accion) {
    case 'login':
      respuesta = procesarLogin(data.usuario, data.password);
      break;
    case 'registrar_asistencia':
      respuesta = registrarAsistencia(data.usuario, data.tipo);
      break;
    case 'crear_empleado':
      respuesta = crearEmpleado(data);
      break;
    case 'listar_empleados':
      respuesta = listarEmpleados();
      break;
    case 'eliminar_empleado':
      respuesta = eliminarEmpleado(data.usuario);
      break;
    case 'listar_asistencias':
      respuesta = listarAsistencias();
      break;
    default:
      respuesta = { exito: false, mensaje: "Acción no reconocida" };
  }

  return ContentService.createTextOutput(JSON.stringify(respuesta))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return doPost(e);
}

// ==================== LOGIN ====================
function procesarLogin(usuario, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Usuarios');
  
  if (!sheet) {
    return { exito: false, mensaje: "No se encontró la hoja 'Usuarios'" };
  }
  
  const datos = sheet.getDataRange().getValues();
  
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === String(usuario).trim() && 
        String(datos[i][1]).trim() === String(password).trim()) {
      return { 
        exito: true, 
        rol: String(datos[i][5]).trim(),
        nombre: datos[i][2],
        area: datos[i][3],
        cargo: datos[i][4]
      };
    }
  }
  return { exito: false, mensaje: "Usuario o contraseña incorrectos" };
}

// ==================== REGISTRAR ASISTENCIA ====================
function registrarAsistencia(usuario, tipo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Asistencias');
  
  if (!sheet) {
    return { exito: false, mensaje: "No se encontró la hoja 'Asistencias'" };
  }
  
  const fecha = new Date();
  sheet.appendRow([usuario, tipo, fecha.toLocaleDateString('es-VE'), fecha.toLocaleTimeString('es-VE'), fecha]);
  return { exito: true, mensaje: `${tipo} registrada a las ${fecha.toLocaleTimeString('es-VE')}` };
}

// ==================== CREAR EMPLEADO ====================
function crearEmpleado(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Usuarios');
  const datos = sheet.getDataRange().getValues();
  
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === String(data.usuario).trim()) {
      return { exito: false, mensaje: "El usuario ya existe" };
    }
  }
  
  sheet.appendRow([data.usuario, data.password, data.nombre, data.area, data.cargo, data.perfil]);
  return { exito: true, mensaje: `Empleado "${data.nombre}" creado exitosamente` };
}

// ==================== LISTAR EMPLEADOS ====================
function listarEmpleados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Usuarios');
  const datos = sheet.getDataRange().getValues();
  const empleados = [];
  
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] && String(datos[i][0]).trim() !== '') {
      empleados.push({
        usuario: String(datos[i][0]).trim(),
        password: String(datos[i][1]).trim(),
        nombre: String(datos[i][2]).trim(),
        area: String(datos[i][3]).trim(),
        cargo: String(datos[i][4]).trim(),
        perfil: String(datos[i][5]).trim()
      });
    }
  }
  
  return { exito: true, empleados: empleados, total: empleados.length };
}

// ==================== ELIMINAR EMPLEADO ====================
function eliminarEmpleado(usuario) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Usuarios');
  const datos = sheet.getDataRange().getValues();
  
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === String(usuario).trim()) {
      sheet.deleteRow(i + 1);
      return { exito: true, mensaje: `Empleado "${usuario}" eliminado` };
    }
  }
  return { exito: false, mensaje: "Empleado no encontrado" };
}

// ==================== LISTAR ASISTENCIAS ====================
function listarAsistencias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Asistencias');
  
  if (!sheet) {
    return { exito: false, mensaje: "No se encontró la hoja 'Asistencias'" };
  }
  
  const datos = sheet.getDataRange().getValues();
  const registros = [];
  
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0]) {
      registros.push({
        usuario: String(datos[i][0]).trim(),
        tipo: String(datos[i][1]).trim(),
        fecha: String(datos[i][2]).trim(),
        hora: String(datos[i][3]).trim()
      });
    }
  }
  
  return { exito: true, registros: registros, total: registros.length };
}