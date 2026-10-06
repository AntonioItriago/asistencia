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
      respuesta = registrarAsistencia(data);
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
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

// ==================== REGISTRAR ASISTENCIA (LÓGICA PRINCIPAL) ====================
function registrarAsistencia(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Asistencias');
  const datos = sheet.getDataRange().getValues();
  const ahora = new Date();
  const fechaHoy = Utilities.formatDate(ahora, 'America/Caracas', 'dd/MM/yyyy');
  const horaAhora = Utilities.formatDate(ahora, 'America/Caracas', 'HH:mm:ss');
  const ubicacion = `${data.lat.toFixed(6)}, ${data.lng.toFixed(6)}`;
  
  // Buscar el registro de HOY para este usuario
  let filaRegistro = -1;
  for (let i = 1; i < datos.length; i++) {
    const usuarioFila = String(datos[i][1]).trim();
    const fechaFila = String(datos[i][3]).trim();
    if (usuarioFila === data.usuario && fechaFila === fechaHoy) {
      filaRegistro = i + 1; // +1 porque las filas en Sheets empiezan en 1
      break;
    }
  }
  
  try {
    switch (data.tipo) {
      case 'entrada':
        // Si ya existe registro hoy, no permitir doble entrada
        if (filaRegistro !== -1) {
          return { exito: false, mensaje: "Ya registraste entrada hoy. Usa los otros botones." };
        }
        
        // Guardar foto en Drive
        let fotoEntradaUrl = '';
        if (data.foto && data.foto.length > 100) {
          fotoEntradaUrl = guardarFoto(data.foto, `${data.usuario}_entrada_${Date.now()}`);
        }
        
        // Generar ID único
        const nuevoId = 'AS-' + ahora.getTime();
        
        // Crear nueva fila: ID(0), Usuario(1), Nombre(2), FechaEntrada(3), HoraEntrada(4), UbicacionEntrada(5), FotoEntrada(6)
        sheet.appendRow([
          nuevoId,
          data.usuario,
          data.nombre || data.usuario,
          fechaHoy,
          horaAhora,
          ubicacion,
          fotoEntradaUrl,
          '', '', '', '', '', '', '', '', '', '', ''
        ]);
        return { exito: true, mensaje: `Entrada registrada a las ${horaAhora}` };
      
      case 'comida-in':
        if (filaRegistro === -1) {
          return { exito: false, mensaje: "Primero debes registrar la Entrada." };
        }
        // Columna H = 8 (HoraIniRef), I = 9 (UbicacionIniRef)
        sheet.getRange(filaRegistro, 8).setValue(horaAhora);
        sheet.getRange(filaRegistro, 9).setValue(ubicacion);
        return { exito: true, mensaje: `Inicio de descanso registrado a las ${horaAhora}` };
      
      case 'comida-out':
        if (filaRegistro === -1) {
          return { exito: false, mensaje: "Primero debes registrar la Entrada." };
        }
        // Verificar que haya comida-in
        const horaIniRef = sheet.getRange(filaRegistro, 8).getValue();
        if (!horaIniRef) {
          return { exito: false, mensaje: "Primero debes registrar Comida-In." };
        }
        // Columna J = 10 (HoraFinRef), K = 11 (UbicacionFinRef)
        sheet.getRange(filaRegistro, 10).setValue(horaAhora);
        sheet.getRange(filaRegistro, 11).setValue(ubicacion);
        return { exito: true, mensaje: `Fin de descanso registrado a las ${horaAhora}` };
      
      case 'salida':
        if (filaRegistro === -1) {
          return { exito: false, mensaje: "Primero debes registrar la Entrada." };
        }
        
        // Guardar foto de salida
        let fotoSalidaUrl = '';
        if (data.foto && data.foto.length > 100) {
          fotoSalidaUrl = guardarFoto(data.foto, `${data.usuario}_salida_${Date.now()}`);
        }
        
        // Columnas: L=12 (FechaSalida), M=13 (HoraSalida), N=14 (UbicacionSalida), O=15 (FotoSalida)
        sheet.getRange(filaRegistro, 12).setValue(fechaHoy);
        sheet.getRange(filaRegistro, 13).setValue(horaAhora);
        sheet.getRange(filaRegistro, 14).setValue(ubicacion);
        sheet.getRange(filaRegistro, 15).setValue(fotoSalidaUrl);
        
        // Calcular tiempos
        const tiempos = calcularTiempos(filaRegistro, sheet);
        sheet.getRange(filaRegistro, 16).setValue(tiempos.productivo); // P
        sheet.getRange(filaRegistro, 17).setValue(tiempos.total);      // Q
        
        return { 
          exito: true, 
          mensaje: `Salida registrada. Tiempo total: ${tiempos.total} | Productivo: ${tiempos.productivo}` 
        };
      
      default:
        return { exito: false, mensaje: "Tipo de registro no válido" };
    }
  } catch (error) {
    return { exito: false, mensaje: "Error: " + error.message };
  }
}

// ==================== CALCULAR TIEMPOS ====================
function calcularTiempos(fila, sheet) {
  const horaEntrada = sheet.getRange(fila, 5).getValue();  // E
  const horaIniRef = sheet.getRange(fila, 8).getValue();   // H
  const horaFinRef = sheet.getRange(fila, 10).getValue();  // J
  const horaSalida = sheet.getRange(fila, 13).getValue();  // M
  
  if (!horaEntrada || !horaSalida) {
    return { productivo: 'N/A', total: 'N/A' };
  }
  
  const entrada = parsearHora(horaEntrada);
  const salida = parsearHora(horaSalida);
  const totalMs = salida - entrada;
  
  let productivoMs = totalMs;
  if (horaIniRef && horaFinRef) {
    const iniRef = parsearHora(horaIniRef);
    const finRef = parsearHora(horaFinRef);
    productivoMs = totalMs - (finRef - iniRef);
  }
  
  return {
    total: formatearDuracion(totalMs),
    productivo: formatearDuracion(productivoMs)
  };
}

function parsearHora(valor) {
  if (valor instanceof Date) return valor;
  const str = String(valor).trim();
  const partes = str.split(':');
  const hoy = new Date();
  hoy.setHours(parseInt(partes[0]), parseInt(partes[1]), parseInt(partes[2]) || 0, 0);
  return hoy;
}

function formatearDuracion(ms) {
  if (ms < 0) return '00:00:00';
  const horas = Math.floor(ms / 3600000);
  const minutos = Math.floor((ms % 3600000) / 60000);
  const segundos = Math.floor((ms % 60000) / 1000);
  return `${String(horas).padStart(2,'0')}:${String(minutos).padStart(2,'0')}:${String(segundos).padStart(2,'0')}`;
}

// ==================== GUARDAR FOTO EN DRIVE ====================
function guardarFoto(base64, nombre) {
  try {
    const carpeta = DriveApp.getFoldersByName('Fotos_Asistencia').hasNext() 
      ? DriveApp.getFoldersByName('Fotos_Asistencia').next() 
      : DriveApp.createFolder('Fotos_Asistencia');
    
    const base64Data = base64.split(',')[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', nombre + '.jpg');
    const archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return archivo.getUrl();
  } catch (error) {
    return 'Error: ' + error.message;
  }
}

// ==================== CREAR EMPLEADO ====================
function crearEmpleado(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Asistencias');
  const datos = sheet.getDataRange().getValues();
  const registros = [];
  
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1]) {
      registros.push({
        id: String(datos[i][0]).trim(),
        usuario: String(datos[i][1]).trim(),
        nombre: String(datos[i][2]).trim(),
        fechaEntrada: String(datos[i][3]).trim(),
        horaEntrada: String(datos[i][4]).trim(),
        ubicacionEntrada: String(datos[i][5]).trim(),
        horaIniRef: String(datos[i][7]).trim(),
        horaFinRef: String(datos[i][9]).trim(),
        fechaSalida: String(datos[i][11]).trim(),
        horaSalida: String(datos[i][12]).trim(),
        tiempoProductivo: String(datos[i][15]).trim(),
        tiempoTotal: String(datos[i][16]).trim()
      });
    }
  }
  
  return { exito: true, registros: registros, total: registros.length };
}