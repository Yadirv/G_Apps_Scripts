// ============================================================
//  NODO 6 — Sistema de Cupones B2B
//  Google Apps Script — Backend completo
//  Hoja: Nueva hoja de cálculo independiente
//
//  HOJAS:
//    - Clientes_B2B       (Módulo Admin)
//    - Cupones            (Módulo Cliente B2B)
//    - Cupones_Reclamados (Módulo Usuario)
//    - Usuarios           (Módulo Usuario)
// ============================================================

// ── Acciones POST ────────────────────────────────────────────
//  add_client | update_client | delete_client
//  add_coupon | update_coupon | delete_coupon
//  register_user | claim_coupon

// ── Acciones GET (query param: action) ──────────────────────
//  get_clients | get_coupons | get_claims | get_all_coupons
//  get_users   | get_stats   | check_user

// ────────────────────────────────────────────────────────────
//  HEADERS DE CADA HOJA
// ────────────────────────────────────────────────────────────
const HEADERS = {
  Clientes_B2B: [
    'Nombre del Negocio', 'Nombre del Contacto', 'C.C o NIT',
    'Celular', 'Correo Electronico', 'Tipo Establecimiento', 'Fecha Registro'
  ],
  Cupones: [
    'Cupon ID', 'CC NIT Cliente', 'Titulo', 'Descripcion',
    'Tipo', 'Categoria', 'Porcentaje Descuento', 'Mascota',
    'Ciudad', 'Fecha Vencimiento', 'Imagen URL',
    'Cantidad Maxima', 'Cantidad Reclamada', 'Estado', 'Fecha Creacion'
  ],
  Cupones_Reclamados: [
    'Reclamacion ID', 'Cupon ID', 'CC NIT Cliente',
    'Nombre Usuario', 'Celular Usuario', 'Ciudad Usuario', 'Mascota Usuario',
    'Fecha Reclamacion', 'Estado'
  ],
  Usuarios: [
    'Celular', 'Nombre Completo', 'Ciudad', 'Mascota',
    'Acepto Politica', 'Fecha Registro', 'Total Cupones Reclamados', 'Cod Tienda'
  ],
  BD_Ciudad: [
    'Ciudad', 'Departamento', 'Activo'
  ],
  Cod_Tiendas: [
    'Nombre del Negocio', 'C.C o NIT', 'Codigo Activo'
  ]
};

// ────────────────────────────────────────────────────────────
//  HELPER: Obtener hoja, crearla y cabecera si no existe
// ────────────────────────────────────────────────────────────
function getSheet(name) {
  var ss = SpreadsheetApp.openById("124qXgkVS93ZMA50PzXngftypXWGX_ZIjOlwIFwvyxhE");
  var sheet = ss.getSheetByName(name);
  var isNew = false;
  if (!sheet) {
    sheet = ss.insertSheet(name);
    isNew = true;
  }
  
  // Si la hoja es nueva o está completamente vacía (como cuando el usuario renombra "Hoja 1")
  if (isNew || sheet.getLastRow() === 0) {
    var headers = HEADERS[name] || [];
    if (headers.length > 0) {
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold')
                 .setBackground('#0f172a')
                 .setFontColor('#ffffff')
                 .setFontSize(10);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// Helper: generar ID único tipo "CUP-A3F9K2"
function generateId(prefix) {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var id = '';
  for (var i = 0; i < 6; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return prefix + '-' + id;
}

// Helper: fecha formateada Colombia
function fechaColombia() {
  return new Date().toLocaleDateString('es-CO', {
    timeZone: 'America/Bogota',
    day: '2-digit', month: '2-digit', year: 'numeric'
  });
}

// Helper: respuesta JSON
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ────────────────────────────────────────────────────────────
//  doGet — Todas las lecturas
// ────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    var action = e.parameter.action;

    // ── GET: Lista de Clientes B2B ─────────────────────────
    if (action === 'get_clients') {
      var sheet = getSheet('Clientes_B2B');
      var rows = sheet.getDataRange().getValues();
      var clientes = [];
      for (var i = 1; i < rows.length; i++) {
        clientes.push({
          negocio:             String(rows[i][0]).trim(),
          contacto:            String(rows[i][1]).trim(),
          cc_nit:              String(rows[i][2]).trim(),
          celular:             String(rows[i][3]).trim(),
          correo:              String(rows[i][4]).trim(),
          tipo_establecimiento:String(rows[i][5]).trim(),
          fecha_registro:      String(rows[i][6]).trim(),
        });
      }
      return jsonResponse(clientes);
    }

    // ── GET: Cupones de un cliente específico ──────────────
    if (action === 'get_coupons') {
      var cc_nit = String(e.parameter.cc_nit || '').trim();
      if (!cc_nit) return jsonResponse({ error: 'cc_nit requerido' });
      var sheet = getSheet('Cupones');
      var rows = sheet.getDataRange().getValues();
      var cupones = [];
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][1]).trim() === cc_nit) {
          cupones.push({
            cupon_id:            String(rows[i][0]).trim(),
            cc_nit_cliente:      String(rows[i][1]).trim(),
            titulo:              String(rows[i][2]).trim(),
            descripcion:         String(rows[i][3]).trim(),
            tipo:                String(rows[i][4]).trim(),
            categoria:           String(rows[i][5]).trim(),
            porcentaje_descuento:Number(rows[i][6]),
            mascota:             String(rows[i][7]).trim(),
            ciudad:              String(rows[i][8]).trim(),
            fecha_vencimiento:   String(rows[i][9]).trim(),
            imagen_url:          String(rows[i][10]).trim(),
            cantidad_maxima:     Number(rows[i][11]),
            cantidad_reclamada:  Number(rows[i][12]),
            estado:              String(rows[i][13]).trim(),
            fecha_creacion:      String(rows[i][14]).trim(),
          });
        }
      }
      return jsonResponse(cupones);
    }

    // ── GET: Reclamaciones recibidas por un cliente ────────
    if (action === 'get_claims') {
      var cc_nit = String(e.parameter.cc_nit || '').trim();
      if (!cc_nit) return jsonResponse({ error: 'cc_nit requerido' });
      var sheet = getSheet('Cupones_Reclamados');
      var rows = sheet.getDataRange().getValues();
      var reclamaciones = [];
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][2]).trim() === cc_nit) {
          reclamaciones.push({
            reclamacion_id:  String(rows[i][0]).trim(),
            cupon_id:        String(rows[i][1]).trim(),
            cc_nit_cliente:  String(rows[i][2]).trim(),
            nombre_usuario:  String(rows[i][3]).trim(),
            celular_usuario: String(rows[i][4]).trim(),
            ciudad_usuario:  String(rows[i][5]).trim(),
            mascota_usuario: String(rows[i][6]).trim(),
            fecha_reclamacion: String(rows[i][7]).trim(),
            estado:          String(rows[i][8]).trim(),
          });
        }
      }
      return jsonResponse(reclamaciones);
    }

    // ── GET: Catálogo público — todos los cupones activos ──
    if (action === 'get_all_coupons') {
      var cod_tienda = String(e.parameter.cod_tienda || '').trim();
      if (!cod_tienda) return jsonResponse({ error: 'Acceso denegado: Código de tienda requerido.' });
      
      var sheetTiendas = getSheet('Cod_Tiendas');
      var rowsTiendas = sheetTiendas.getDataRange().getValues();
      var codigoValido = false;
      for (var i = 1; i < rowsTiendas.length; i++) {
        if (String(rowsTiendas[i][2]).trim() === cod_tienda) {
          codigoValido = true;
          break;
        }
      }
      if (!codigoValido) return jsonResponse({ error: 'Acceso denegado: Código de tienda inválido.' });

      var sheetC = getSheet('Cupones');
      var sheetB = getSheet('Clientes_B2B');
      var rowsC  = sheetC.getDataRange().getValues();
      var rowsB  = sheetB.getDataRange().getValues();
      var hoy    = new Date();
      hoy.setHours(0,0,0,0);

      // Mapa cc_nit → nombre del negocio
      var negocios = {};
      for (var i = 1; i < rowsB.length; i++) {
        negocios[String(rowsB[i][2]).trim()] = {
          negocio: String(rowsB[i][0]).trim(),
          celular: String(rowsB[i][3]).trim(),
        };
      }

      var cupones = [];
      for (var i = 1; i < rowsC.length; i++) {
        var estadoActual = String(rowsC[i][13]).trim();
        var fechaVence   = new Date(String(rowsC[i][9]).trim());
        fechaVences      = isNaN(fechaVence) ? null : fechaVence;

        // Auto-vencer si la fecha pasó
        if (fechaVences && fechaVences < hoy) estadoActual = 'Vencido';
        if (estadoActual !== 'Activo') continue;

        var ccNit = String(rowsC[i][1]).trim();
        var negInfo = negocios[ccNit] || { negocio: ccNit, celular: '' };

        cupones.push({
          cupon_id:            String(rowsC[i][0]).trim(),
          cc_nit_cliente:      ccNit,
          nombre_negocio:      negInfo.negocio,
          celular_negocio:     negInfo.celular,
          titulo:              String(rowsC[i][2]).trim(),
          descripcion:         String(rowsC[i][3]).trim(),
          tipo:                String(rowsC[i][4]).trim(),
          categoria:           String(rowsC[i][5]).trim(),
          porcentaje_descuento:Number(rowsC[i][6]),
          mascota:             String(rowsC[i][7]).trim(),
          ciudad:              String(rowsC[i][8]).trim(),
          fecha_vencimiento:   String(rowsC[i][9]).trim(),
          imagen_url:          String(rowsC[i][10]).trim(),
          cantidad_maxima:     Number(rowsC[i][11]),
          cantidad_reclamada:  Number(rowsC[i][12]),
          estado:              estadoActual,
          fecha_creacion:      String(rowsC[i][14]).trim(),
          row_index:           i + 1, // fila real en la hoja (para actualizar)
        });
      }
      return jsonResponse(cupones);
    }

    // ── GET: Todos los usuarios ────────────────────────────
    if (action === 'get_users') {
      var sheet = getSheet('Usuarios');
      var rows = sheet.getDataRange().getValues();
      var usuarios = [];
      for (var i = 1; i < rows.length; i++) {
        usuarios.push({
          celular:                  String(rows[i][0]).trim(),
          nombre_completo:          String(rows[i][1]).trim(),
          ciudad:                   String(rows[i][2]).trim(),
          mascota:                  String(rows[i][3]).trim(),
          acepto_politica:          String(rows[i][4]).trim(),
          fecha_registro:           String(rows[i][5]).trim(),
          total_cupones_reclamados: Number(rows[i][6]),
        });
      }
      return jsonResponse(usuarios);
    }

    // ── GET: Verificar si usuario ya existe ────────────────
    if (action === 'check_user') {
      var celular = String(e.parameter.celular || '').trim();
      if (!celular) return jsonResponse({ exists: false });
      var sheet = getSheet('Usuarios');
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === celular) {
          return jsonResponse({
            exists: true,
            nombre_completo: String(rows[i][1]).trim(),
            ciudad:          String(rows[i][2]).trim(),
            mascota:         String(rows[i][3]).trim(),
            cod_tienda:      String(rows[i][7] || '').trim()
          });
        }
      }
      return jsonResponse({ exists: false });
    }

    // ── GET: Métricas para Dashboard Admin ─────────────────
    if (action === 'get_stats') {
      var sheetU = getSheet('Usuarios');
      var sheetR = getSheet('Cupones_Reclamados');
      var sheetC = getSheet('Cupones');

      var rowsU = sheetU.getDataRange().getValues();
      var rowsR = sheetR.getDataRange().getValues();
      var rowsC = sheetC.getDataRange().getValues();

      // Usuarios por ciudad
      var porCiudad = {};
      // Usuarios por mascota
      var porMascota = { 'Perro': 0, 'Gato': 0, 'Ambos': 0 };
      // Registros por semana (últimas 8 semanas)
      var porSemana = {};

      for (var i = 1; i < rowsU.length; i++) {
        var ciudad  = String(rowsU[i][2]).trim() || 'Sin ciudad';
        var mascota = String(rowsU[i][3]).trim();
        porCiudad[ciudad] = (porCiudad[ciudad] || 0) + 1;
        if (porMascota[mascota] !== undefined) porMascota[mascota]++;
      }

      // Top 10 cupones más reclamados
      var recPorCupon = {};
      for (var i = 1; i < rowsR.length; i++) {
        var cid = String(rowsR[i][1]).trim();
        recPorCupon[cid] = (recPorCupon[cid] || 0) + 1;
      }
      // Unir con títulos
      var cuponTitulos = {};
      for (var i = 1; i < rowsC.length; i++) {
        cuponTitulos[String(rowsC[i][0]).trim()] = String(rowsC[i][2]).trim();
      }
      var topCupones = Object.entries(recPorCupon)
        .sort(function(a,b) { return b[1] - a[1]; })
        .slice(0,10)
        .map(function(item) {
          var id = item[0], count = item[1];
          return { id: id, titulo: cuponTitulos[id] || id, count: count };
        });

      return jsonResponse({
        total_usuarios:           Math.max(0, rowsU.length - 1),
        total_reclamaciones:      Math.max(0, rowsR.length - 1),
        total_cupones:            Math.max(0, rowsC.length - 1),
        por_ciudad:               porCiudad,
        por_mascota:              porMascota,
        top_cupones:              topCupones,
      });
    }

    // ── GET: Lista de Ciudades desde BD_Ciudad ────────────────
    if (action === 'get_ciudades') {
      var sheet = getSheet('BD_Ciudad');
      var rows = sheet.getDataRange().getValues();
      var ciudades = [];
      for (var i = 1; i < rows.length; i++) {
        var ciudad  = String(rows[i][0]).trim();
        var activo  = String(rows[i][2]).trim().toLowerCase();
        // Solo incluir si está marcado como Activo (o si la columna está vacía)
        if (ciudad && activo !== 'no' && activo !== 'false' && activo !== '0') {
          ciudades.push({
            ciudad:       ciudad,
            departamento: String(rows[i][1]).trim(),
          });
        }
      }
      // Ordenar alfabéticamente
      ciudades.sort(function(a,b){ return a.ciudad.localeCompare(b.ciudad, 'es'); });
      return jsonResponse(ciudades);
    }

    return jsonResponse({ error: 'Accion GET no reconocida: ' + action });

  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

// ────────────────────────────────────────────────────────────
//  doPost — Todas las escrituras
// ────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var action = data.action;

    // ── POST: Crear Cliente B2B ────────────────────────────
    if (action === 'add_client') {
      var sheet = getSheet('Clientes_B2B');
      var rows  = sheet.getDataRange().getValues();

      // Validar duplicado por CC/NIT o Celular
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][2]).trim() === String(data.cc_nit).trim() ||
            String(rows[i][3]).trim() === String(data.celular).trim()) {
          return jsonResponse({ status: 'error', message: 'Ya existe un cliente con ese C.C/NIT o Celular.' });
        }
      }
      sheet.appendRow([
        data.negocio, data.contacto, data.cc_nit,
        data.celular, data.correo, data.tipo_establecimiento,
        fechaColombia()
      ]);
      return jsonResponse({ status: 'success', message: 'Cliente registrado exitosamente.' });
    }

    // ── POST: Actualizar Cliente B2B ───────────────────────
    if (action === 'update_client') {
      var sheet = getSheet('Clientes_B2B');
      var rows  = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][2]).trim() === String(data.cc_nit).trim()) {
          var rowNum = i + 1;
          sheet.getRange(rowNum, 1, 1, 7).setValues([[
            data.negocio, data.contacto, data.cc_nit,
            data.celular, data.correo, data.tipo_establecimiento,
            String(rows[i][6]) // conservar fecha registro original
          ]]);
          return jsonResponse({ status: 'success', message: 'Cliente actualizado.' });
        }
      }
      return jsonResponse({ status: 'error', message: 'Cliente no encontrado.' });
    }

    // ── POST: Eliminar Cliente B2B ─────────────────────────
    if (action === 'delete_client') {
      var sheet = getSheet('Clientes_B2B');
      var rows  = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][2]).trim() === String(data.cc_nit).trim()) {
          sheet.deleteRow(i + 1);
          return jsonResponse({ status: 'success', message: 'Cliente eliminado.' });
        }
      }
      return jsonResponse({ status: 'error', message: 'Cliente no encontrado.' });
    }

    // ── POST: Crear Cupón ──────────────────────────────────
    if (action === 'add_coupon') {
      var sheet  = getSheet('Cupones');
      var cuponId = generateId('CUP');
      sheet.appendRow([
        cuponId,
        data.cc_nit_cliente,
        data.titulo,
        data.descripcion     || '',
        data.tipo,
        data.categoria,
        Number(data.porcentaje_descuento),
        data.mascota,
        data.ciudad,
        data.fecha_vencimiento,
        data.imagen_url      || '',
        Number(data.cantidad_maxima),
        0,                    // cantidad_reclamada inicial = 0
        'Activo',
        fechaColombia()
      ]);
      return jsonResponse({ status: 'success', message: 'Cupón creado.', cupon_id: cuponId });
    }

    // ── POST: Actualizar Cupón ─────────────────────────────
    if (action === 'update_coupon') {
      var sheet = getSheet('Cupones');
      var rows  = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(data.cupon_id).trim() &&
            String(rows[i][1]).trim() === String(data.cc_nit_cliente).trim()) {
          // Gate: no actualizar si tiene reclamaciones con estado Reclamado
          var cantRec = Number(rows[i][12]);
          if (cantRec > 0) {
            // Solo actualizar campos no críticos (descripción, imagen, cantidad máxima)
            var rowNum = i + 1;
            sheet.getRange(rowNum, 3, 1, 9).setValues([[
              data.titulo,
              data.descripcion || String(rows[i][3]),
              data.tipo        || String(rows[i][4]),
              data.categoria   || String(rows[i][5]),
              Number(data.porcentaje_descuento) || Number(rows[i][6]),
              data.mascota     || String(rows[i][7]),
              data.ciudad      || String(rows[i][8]),
              data.fecha_vencimiento || String(rows[i][9]),
              data.imagen_url  || String(rows[i][10]),
            ]]);
            // Actualizar cantidad_maxima (col 12)
            sheet.getRange(rowNum, 12).setValue(Number(data.cantidad_maxima) || Number(rows[i][11]));
          } else {
            var rowNum = i + 1;
            sheet.getRange(rowNum, 3, 1, 9).setValues([[
              data.titulo, data.descripcion || '', data.tipo,
              data.categoria, Number(data.porcentaje_descuento),
              data.mascota, data.ciudad, data.fecha_vencimiento,
              data.imagen_url || ''
            ]]);
            sheet.getRange(rowNum, 12).setValue(Number(data.cantidad_maxima));
          }
          return jsonResponse({ status: 'success', message: 'Cupón actualizado.' });
        }
      }
      return jsonResponse({ status: 'error', message: 'Cupón no encontrado.' });
    }

    // ── POST: Eliminar Cupón ───────────────────────────────
    if (action === 'delete_coupon') {
      // Gate de seguridad: verificar que no tenga reclamaciones activas
      var sheetR = getSheet('Cupones_Reclamados');
      var rowsR  = sheetR.getDataRange().getValues();
      for (var i = 1; i < rowsR.length; i++) {
        if (String(rowsR[i][1]).trim() === String(data.cupon_id).trim() &&
            String(rowsR[i][8]).trim() === 'Reclamado') {
          return jsonResponse({
            status: 'error',
            message: 'No se puede eliminar: el cupón tiene reclamaciones activas.'
          });
        }
      }
      var sheetC = getSheet('Cupones');
      var rowsC  = sheetC.getDataRange().getValues();
      for (var i = 1; i < rowsC.length; i++) {
        if (String(rowsC[i][0]).trim() === String(data.cupon_id).trim() &&
            String(rowsC[i][1]).trim() === String(data.cc_nit_cliente).trim()) {
          sheetC.deleteRow(i + 1);
          return jsonResponse({ status: 'success', message: 'Cupón eliminado.' });
        }
      }
      return jsonResponse({ status: 'error', message: 'Cupón no encontrado.' });
    }

    // ── POST: Registrar Usuario ────────────────────────────
    if (action === 'register_user') {
      var cod_tienda = String(data.cod_tienda || '').trim();
      if (!cod_tienda) {
        return jsonResponse({ status: 'error', message: 'Debes ingresar el código de la tienda.' });
      }

      // Validar código contra hoja Cod_Tiendas
      var sheetTiendas = getSheet('Cod_Tiendas');
      var rowsTiendas = sheetTiendas.getDataRange().getValues();
      var codigoValido = false;
      for (var i = 1; i < rowsTiendas.length; i++) {
        if (String(rowsTiendas[i][2]).trim() === cod_tienda) {
          codigoValido = true;
          break;
        }
      }

      if (!codigoValido) {
        return jsonResponse({ status: 'error', message: 'Código incorrecto o inactivo. Solicítalo en caja al realizar tu compra.' });
      }

      var sheet  = getSheet('Usuarios');
      var rows   = sheet.getDataRange().getValues();
      var celular = String(data.celular).trim();

      // Si ya existe, actualizar datos (idempotente)
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === celular) {
          sheet.getRange(i + 1, 2, 1, 3).setValues([[
            data.nombre_completo, data.ciudad, data.mascota
          ]]);
          // Actualizar cod_tienda en la columna 8 (índice 7)
          sheet.getRange(i + 1, 8).setValue(cod_tienda);
          return jsonResponse({ status: 'success', message: 'Usuario actualizado.', existing: true });
        }
      }
      // Nuevo usuario
      sheet.appendRow([
        celular, data.nombre_completo, data.ciudad, data.mascota,
        'SÍ', fechaColombia(), 0, cod_tienda
      ]);
      return jsonResponse({ status: 'success', message: 'Usuario registrado.', existing: false });
    }

    // ── POST: Reclamar Cupón ───────────────────────────────
    if (action === 'claim_coupon') {
      var sheetC  = getSheet('Cupones');
      var sheetR  = getSheet('Cupones_Reclamados');
      var sheetU  = getSheet('Usuarios');
      var rowsC   = sheetC.getDataRange().getValues();
      var rowsR   = sheetR.getDataRange().getValues();
      var rowsU   = sheetU.getDataRange().getValues();
      var cuponId = String(data.cupon_id).trim();
      var celular = String(data.celular_usuario).trim();
      var hoy     = new Date();
      hoy.setHours(0,0,0,0);

      // Buscar el cupón
      var cuponRow = null;
      var cuponRowNum = -1;
      for (var i = 1; i < rowsC.length; i++) {
        if (String(rowsC[i][0]).trim() === cuponId) {
          cuponRow    = rowsC[i];
          cuponRowNum = i + 1;
          break;
        }
      }
      if (!cuponRow) return jsonResponse({ status: 'error', message: 'Cupón no encontrado.' });

      // Validar estado activo
      if (String(cuponRow[13]).trim() !== 'Activo') {
        return jsonResponse({ status: 'error', message: 'Este cupón ya no está activo.' });
      }

      // Validar vencimiento
      var fechaVence = new Date(String(cuponRow[9]).trim());
      if (!isNaN(fechaVence) && fechaVence < hoy) {
        // Marcar como vencido
        sheetC.getRange(cuponRowNum, 14).setValue('Vencido');
        return jsonResponse({ status: 'error', message: 'Este cupón está vencido.' });
      }

      // Validar disponibilidad
      var cantMax = Number(cuponRow[11]);
      var cantRec = Number(cuponRow[12]);
      if (cantRec >= cantMax) {
        sheetC.getRange(cuponRowNum, 14).setValue('Agotado');
        return jsonResponse({ status: 'error', message: 'Este cupón ya fue agotado.' });
      }

      // Validar que el usuario no haya reclamado YA este cupón
      for (var i = 1; i < rowsR.length; i++) {
        if (String(rowsR[i][1]).trim() === cuponId &&
            String(rowsR[i][4]).trim() === celular) {
          return jsonResponse({ status: 'error', message: 'Ya reclamaste este cupón anteriormente.' });
        }
      }

      // Crear reclamación
      var recId = generateId('REC');
      sheetR.appendRow([
        recId, cuponId, String(cuponRow[1]).trim(),
        data.nombre_usuario, celular, data.ciudad_usuario, data.mascota_usuario,
        fechaColombia(), 'Reclamado'
      ]);

      // Incrementar cantidad_reclamada
      var nuevaCantRec = cantRec + 1;
      sheetC.getRange(cuponRowNum, 13).setValue(nuevaCantRec);

      // Cambiar estado a Agotado si se alcanzó el máximo
      if (nuevaCantRec >= cantMax) {
        sheetC.getRange(cuponRowNum, 14).setValue('Agotado');
      }

      // Actualizar contador del usuario
      for (var i = 1; i < rowsU.length; i++) {
        if (String(rowsU[i][0]).trim() === celular) {
          sheetU.getRange(i + 1, 7).setValue(Number(rowsU[i][6]) + 1);
          break;
        }
      }

      // Obtener datos del Cliente B2B para la notificación
      var sheetB = getSheet('Clientes_B2B');
      var rowsB  = sheetB.getDataRange().getValues();
      var negocioInfo = { negocio: '', celular: '' };
      for (var i = 1; i < rowsB.length; i++) {
        if (String(rowsB[i][2]).trim() === String(cuponRow[1]).trim()) {
          negocioInfo = { negocio: String(rowsB[i][0]).trim(), celular: String(rowsB[i][3]).trim() };
          break;
        }
      }

      // Retornar todos los datos necesarios para las notificaciones WhatsApp
      return jsonResponse({
        status:   'success',
        message:  'Cupón reclamado exitosamente.',
        reclamacion_id: recId,
        // Datos para notificación al usuario
        notif_usuario: {
          celular:       celular,
          nombre:        data.nombre_usuario,
          titulo_cupon:  String(cuponRow[2]).trim(),
          nombre_negocio: negocioInfo.negocio,
          celular_negocio: negocioInfo.celular,
        },
        // Datos para notificación al Cliente B2B
        notif_negocio: {
          celular:            negocioInfo.celular,
          nombre_negocio:     negocioInfo.negocio,
          cupon_id:           cuponId,
          titulo_cupon:       String(cuponRow[2]).trim(),
          porcentaje:         Number(cuponRow[6]),
          nombre_usuario:     data.nombre_usuario,
          celular_usuario:    celular,
          ciudad_usuario:     data.ciudad_usuario,
        }
      });
    }

    return jsonResponse({ status: 'error', message: 'Accion no reconocida: ' + action });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}





