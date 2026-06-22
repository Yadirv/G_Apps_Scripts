// ==========================================
// GOOGLE APPS SCRIPT: RECEPTOR DE PEDIDOS B2B Y CLIENTES
// ==========================================

function doPost(e) {
  // Configurar cabeceras CORS
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;

    // ------------------------------------------
    // ACCIÓN: AÑADIR NUEVO CLIENTE (DESDE ADMIN)
    // ------------------------------------------
    if (action === "add_client") {
      var sheetClientes = SpreadsheetApp.openById("1ZHrSO_17gxhP0ZRL_lKsJPKmk9Fl2mv9tP_JNvIYWfc").getSheetByName("Clientes");
      
      if (!sheetClientes) {
         sheetClientes = SpreadsheetApp.openById("1ZHrSO_17gxhP0ZRL_lKsJPKmk9Fl2mv9tP_JNvIYWfc").insertSheet("Clientes");
      }
      
      // Si la hoja está totalmente limpia, inicializar fila de títulos
      if (sheetClientes.getLastRow() === 0) {
        sheetClientes.appendRow(["Nombre del Negocio", "Nombre del Contacto", "C.C o NIT", "Celular", "Correo", "Canales"]);
        sheetClientes.getRange("A1:F1").setFontWeight("bold").setBackground("#1E293B").setFontColor("#FFFFFF");
      }

      // Validar que no exista ya el cliente (por CC/NIT o Celular)
      var datosActuales = sheetClientes.getDataRange().getValues();
      for (var i = 1; i < datosActuales.length; i++) {
        var ccNitExistente = String(datosActuales[i][2]).trim();
        var celularExistente = String(datosActuales[i][3]).trim();
        if (ccNitExistente === String(data.cc_nit).trim() || celularExistente === String(data.celular).trim()) {
           return ContentService.createTextOutput(JSON.stringify({ 
             "status": "error", 
             "message": "El cliente ya existe con ese C.C/NIT o Celular." 
           })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      sheetClientes.appendRow([
        data.negocio,
        data.contacto,
        data.cc_nit,
        data.celular,
        data.correo,
        data.canales || ""
      ]);

      return ContentService.createTextOutput(JSON.stringify({ 
        "status": "success", 
        "message": "Cliente registrado exitosamente" 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ------------------------------------------
    // ACCIÓN: ACTUALIZAR CLIENTE EXISTENTE
    // ------------------------------------------
    if (action === "update_client") {
      var sheetClientes = SpreadsheetApp.openById("1ZHrSO_17gxhP0ZRL_lKsJPKmk9Fl2mv9tP_JNvIYWfc").getSheetByName("Clientes");
      if (!sheetClientes) {
        throw new Error("La hoja de Clientes no existe.");
      }

      var ccNitBuscado = String(data.cc_nit).trim();
      if (!ccNitBuscado) {
        throw new Error("Se requiere C.C/NIT para actualizar un cliente.");
      }

      var datosActuales = sheetClientes.getDataRange().getValues();
      var filaEncontrada = -1;

      // Iterar desde fila 1 (asumiendo fila 0 son títulos)
      for (var i = 1; i < datosActuales.length; i++) {
        var ccNitExistente = String(datosActuales[i][2]).trim();
        if (ccNitExistente === ccNitBuscado) {
          filaEncontrada = i + 1; // +1 porque la fila en Sheet es 1-indexed
          break;
        }
      }

      if (filaEncontrada !== -1) {
        // Actualizar valores en la hoja: [Negocio(A), Contacto(B), CC(C), Celular(D), Correo(E), Canales(F)]
        if (data.negocio) sheetClientes.getRange(filaEncontrada, 1).setValue(data.negocio);
        if (data.contacto) sheetClientes.getRange(filaEncontrada, 2).setValue(data.contacto);
        if (data.celular) sheetClientes.getRange(filaEncontrada, 4).setValue(data.celular);
        if (data.correo) sheetClientes.getRange(filaEncontrada, 5).setValue(data.correo);
        if (data.canales !== undefined) sheetClientes.getRange(filaEncontrada, 6).setValue(data.canales);

        return ContentService.createTextOutput(JSON.stringify({ 
          "status": "success", 
          "message": "Cliente actualizado exitosamente" 
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ 
          "status": "error", 
          "message": "No se encontró el cliente con CC/NIT: " + ccNitBuscado
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ------------------------------------------
    // ACCIÓN: ELIMINAR CLIENTE
    // ------------------------------------------
    if (action === "delete_client") {
      var sheetClientes = SpreadsheetApp.openById("1ZHrSO_17gxhP0ZRL_lKsJPKmk9Fl2mv9tP_JNvIYWfc").getSheetByName("Clientes");
      var dataClientes = sheetClientes.getDataRange().getValues();
      var ccNitBuscado = String(data.cc_nit).trim();
      
      var filaEncontrada = -1;
      
      // Buscar el cliente por CC/NIT (Columna 3)
      for (var i = 1; i < dataClientes.length; i++) {
        var nitActual = String(dataClientes[i][2]).trim();
        if (nitActual === ccNitBuscado) {
          filaEncontrada = i + 1; // +1 porque getValues es 0-indexed y sheets son 1-indexed
          break;
        }
      }
      
      if (filaEncontrada !== -1) {
        sheetClientes.deleteRow(filaEncontrada);
        return ContentService.createTextOutput(JSON.stringify({ 
          "status": "success", 
          "message": "Cliente eliminado exitosamente" 
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ 
          "status": "error", 
          "message": "No se encontró el cliente con CC/NIT: " + ccNitBuscado
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ------------------------------------------
    // ACCIÓN: REGISTRAR PEDIDO (DESDE DASHBOARD)
    // ------------------------------------------
    if (action === "add_order") {
      var sheetPedidos = SpreadsheetApp.openById("1ZHrSO_17gxhP0ZRL_lKsJPKmk9Fl2mv9tP_JNvIYWfc").getSheetByName("Pedidos Recibidos");
      
      if (!sheetPedidos) {
         sheetPedidos = SpreadsheetApp.openById("1ZHrSO_17gxhP0ZRL_lKsJPKmk9Fl2mv9tP_JNvIYWfc").insertSheet("Pedidos Recibidos");
      }

      // Si la hoja está totalmente limpia, inicializar fila de títulos
      if (sheetPedidos.getLastRow() === 0) {
        sheetPedidos.appendRow(["Fecha", "CC/NIT", "Celular", "Nombre Negocio", "Referencia", "Producto", "Marca", "Cantidad", "Precio Unitario", "Descuento Aplicado", "Aplica IVA", "Total Item"]);
        sheetPedidos.getRange("A1:L1").setFontWeight("bold").setBackground("#1E293B").setFontColor("#FFFFFF");
      }

      var cc_nit = data.cc_nit;
      var celular = data.celular;
      var negocio = data.negocio;
      var fecha = data.fecha;
      var incluyeIva = data.incluye_iva ? "SÍ" : "NO";
      var totalPuntos = data.totalPuntos || 0;
      var items = data.items;

      // Iterar e insertar de manera incremental cada ítem para evitar sobreescritura
      items.forEach(function(item) {
        sheetPedidos.appendRow([
          fecha,
          cc_nit,
          celular,
          negocio,
          item.referencia,
          item.producto,
          item.marca,
          item.cantidad,
          item.precioUnitario,
          item.descuentoAplicado,
          incluyeIva,
          item.totalItem
        ]);
      });

      // LÓGICA DE PUNTOS
      if (totalPuntos > 0) {
        var sheetPuntos = SpreadsheetApp.openById("1ZHrSO_17gxhP0ZRL_lKsJPKmk9Fl2mv9tP_JNvIYWfc").getSheetByName("Puntos_Clientes");
        if (!sheetPuntos) {
          sheetPuntos = SpreadsheetApp.openById("1ZHrSO_17gxhP0ZRL_lKsJPKmk9Fl2mv9tP_JNvIYWfc").insertSheet("Puntos_Clientes");
        }
        if (sheetPuntos.getLastRow() === 0) {
          sheetPuntos.appendRow(["C.C o NIT", "Nombre del Negocio", "Fecha", "Total Puntos"]);
          sheetPuntos.getRange("A1:D1").setFontWeight("bold").setBackground("#10B981").setFontColor("#FFFFFF");
        }
        sheetPuntos.appendRow([cc_nit, negocio, fecha, totalPuntos]);
      }

      return ContentService.createTextOutput(JSON.stringify({ 
        "status": "success", 
        "message": "Pedido consolidado exitosamente" 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    throw new Error("Acción no válida");

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      "status": "error", 
      "message": error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ------------------------------------------
// GOOGLE APPS SCRIPT: RESPONDER A PREFLIGHT (CORS) Y GET
// ------------------------------------------
function doOptions(e) {
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
}

function doGet(e) {
  try {
    var sheetClientes = SpreadsheetApp.openById("1ZHrSO_17gxhP0ZRL_lKsJPKmk9Fl2mv9tP_JNvIYWfc").getSheetByName("Clientes");
    
    // Si no existe, retornamos array vacío
    if (!sheetClientes || sheetClientes.getLastRow() <= 1) {
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var datos = sheetClientes.getDataRange().getValues();
    var clientes = [];
    
    // Fila 1 = Títulos: [Negocio, Contacto, CC/NIT, Celular, Correo, Canales]
    for (var i = 1; i < datos.length; i++) {
      clientes.push({
        negocio: String(datos[i][0]).trim(),
        contacto: String(datos[i][1]).trim(),
        cc_nit: String(datos[i][2]).trim(),
        celular: String(datos[i][3]).trim(),
        correo: String(datos[i][4]).trim(),
        canales: String(datos[i][5] || "").trim()
      });
    }

    return ContentService.createTextOutput(JSON.stringify(clientes))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "error": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}



