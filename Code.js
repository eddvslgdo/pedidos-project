function doGet() {
    return HtmlService.createTemplateFromFile('modal')
        .evaluate()
        .setTitle('Gestor Integral de Pedidos - Polaquimia')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obtenerMetricasKPIs() {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPedidos = libro.getSheetByName('MONITOREO DE PEDIDOS');
    const sheetProductos = libro.getSheetByName('productos');

    if (!sheetPedidos) return { error: "No se encontró la pestaña 'MONITOREO DE PEDIDOS'" };

    // --- LECTURA DE PEDIDOS ---
    const datosPedidos = sheetPedidos.getDataRange().getValues();
    const cabecerasPedidos = datosPedidos[0].map(c => String(c).trim().toUpperCase());
    const filasPedidos = datosPedidos.slice(1);

    const getIdx = (nombre) => cabecerasPedidos.indexOf(nombre);

    const idxOC = getIdx('OC CLIENTE');
    const idxPedido = getIdx('PEDIDO CLIENTE');
    const idxCreacion = getIdx('FECHA CREACIÓN PEDIDO');
    const idxPedidoInter = getIdx('PEDIDO INTER');
    const idxOcInter = getIdx('OC INTER');
    const idxOrg = getIdx('ORG');
    const idxCliente = getIdx('CLIENTE');
    const idxMaterial = getIdx('MATERIAL');
    const idxSkuPq = getIdx('SKU PQ');
    const idxProducto = getIdx('PRODUCTO');
    const idxAlias = getIdx('ALIAS');
    const idxCantidad = getIdx('CANTIDAD');
    const idxUm = getIdx('UM');
    const idxProg = getIdx('ENTREGA PROGRAMADA');
    const idxEntrega = getIdx('FECHA DE ENTREGA');
    const idxEstatus = getIdx('ESTATUS');
    const idxEstatusTiempo = getIdx('ESTATUS DE TIEMPO');
    const idxOrigen = getIdx('ORIGEN');
    const idxEnvio = getIdx('TIPO DE ENVIO') !== -1 ? getIdx('TIPO DE ENVIO') : getIdx('TIPO DE ENVÍO');
    const idxReqInter = getIdx('PEDIDO INTER REQUERIDO');

    let ocsTotales = 0, pedidosConvertidos = 0;
    const historialPedidos = [];
    const clientesUnicos = new Set();
    const orgsUnicas = new Set();
    const materialesUnicos = new Set();

    const formatearFechaStr = (fecha) => {
        if (!fecha || fecha === "") return "N/A";
        let d = new Date(fecha);
        return isNaN(d.getTime()) ? "N/A" : d.toLocaleDateString('es-MX');
    };

    const formatoISO = (fecha) => {
        if (!fecha || fecha === "") return null;
        let d = new Date(fecha);
        return isNaN(d.getTime()) ? null : d.toISOString().split('T')[0];
    };

    filasPedidos.forEach((fila, nFila) => {
        const oc = idxOC !== -1 ? String(fila[idxOC]) : "";
        const pedido = idxPedido !== -1 ? String(fila[idxPedido]) : "";
        const pedInter = idxPedidoInter !== -1 ? String(fila[idxPedidoInter]) : "";
        const ocInter = idxOcInter !== -1 ? String(fila[idxOcInter]) : "";

        const org = idxOrg !== -1 ? String(fila[idxOrg]) : "";
        const cliente = idxCliente !== -1 ? String(fila[idxCliente]) : "";
        const material = idxMaterial !== -1 ? String(fila[idxMaterial]) : "";

        const skuPq = idxSkuPq !== -1 ? String(fila[idxSkuPq]) : "";
        const productoNombre = idxProducto !== -1 ? String(fila[idxProducto]) : "";
        const alias = idxAlias !== -1 ? String(fila[idxAlias]) : "";

        const cantidad = idxCantidad !== -1 ? (parseFloat(fila[idxCantidad]) || 0) : 0;
        const um = idxUm !== -1 ? String(fila[idxUm]) : "";

        const estatus = idxEstatus !== -1 ? String(fila[idxEstatus]) : "";
        const estatusTiempo = idxEstatusTiempo !== -1 ? String(fila[idxEstatusTiempo]) : "";

        if (!oc && !pedido && !pedInter && !ocInter) return;

        if (oc || ocInter) ocsTotales++;
        if (pedido || pedInter) {
            pedidosConvertidos++;
            if (cliente) clientesUnicos.add(cliente);
            if (org) orgsUnicas.add(org);
            if (material) materialesUnicos.add(material);
        }

        let estadoEntrega = "En Proceso";
        if (estatus === "Completado") {
            estadoEntrega = estatusTiempo ? estatusTiempo : "Completado";
        } else if (estatus === "Cancelado") {
            estadoEntrega = "Cancelado";
        } else if (estatus === "En Tránsito") {
            estadoEntrega = "En Tránsito";
        }

        historialPedidos.push({
            filaId: nFila + 2,
            oc, pedido,
            pedInter: pedInter || "—",
            ocInter: ocInter || "—",
            cliente, org, material, cantidad, um,
            productoNombre: productoNombre,
            productoSKUPQ: skuPq,
            alias: alias,
            estadoEntrega,
            reqInter: idxReqInter !== -1 ? fila[idxReqInter] : "No",
            origen: idxOrigen !== -1 ? fila[idxOrigen] : "",
            tipoEnvio: idxEnvio !== -1 ? fila[idxEnvio] : "",
            fechaCreacionISO: formatoISO(idxCreacion !== -1 ? fila[idxCreacion] : ""),
            fechaCreacion: formatearFechaStr(idxCreacion !== -1 ? fila[idxCreacion] : ""),
            fechaProg: formatearFechaStr(idxProg !== -1 ? fila[idxProg] : ""),
            fechaEntrega: formatearFechaStr(idxEntrega !== -1 ? fila[idxEntrega] : ""),
            fechaProgRaw: formatoISO(idxProg !== -1 ? fila[idxProg] : "")
        });
    });

    // --- LECTURA SEGURA DE CATÁLOGO DE PRODUCTOS ---
    let catalogoProductos = {};
    if (sheetProductos) {
        const datosProd = sheetProductos.getDataRange().getValues();
        if (datosProd.length > 1) {
            const cabProd = datosProd[0].map(c => String(c).trim().toLowerCase());
            const filasProd = datosProd.slice(1);

            const idxSkuDjpProd = cabProd.indexOf('sku djp');
            const idxSkuPqProd = cabProd.indexOf('sku pq');
            const idxNombreProd = cabProd.indexOf('nombre del producto');
            const idxAliasProd = cabProd.indexOf('alias');

            filasProd.forEach(fila => {
                const nombre = idxNombreProd !== -1 ? String(fila[idxNombreProd]).trim() : "";
                // Solo procesamos la fila si tiene un nombre (ignoramos filas vacías)
                if (nombre && nombre !== "") {
                    catalogoProductos[nombre] = {
                        skuPq: idxSkuPqProd !== -1 ? String(fila[idxSkuPqProd]).trim() : "",
                        skuDjp: idxSkuDjpProd !== -1 ? String(fila[idxSkuDjpProd]).trim() : "",
                        alias: idxAliasProd !== -1 ? String(fila[idxAliasProd]).trim() : ""
                    };
                }
            });
        }
    }

    return {
        tasaConversionOC: ocsTotales > 0 ? ((pedidosConvertidos / ocsTotales) * 100).toFixed(1) : 0,
        historialPedidos,
        clientesLista: Array.from(clientesUnicos),
        orgsLista: Array.from(orgsUnicas),
        materialesLista: Array.from(materialesUnicos),
        catalogoProductos: catalogoProductos
    };
}

function guardarNuevoPedido(objetoPedido) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MONITOREO DE PEDIDOS');
    const cabeceras = sheet.getDataRange().getValues()[0].map(c => String(c).trim().toUpperCase());

    let numPedInter = objetoPedido.pedidoInter;
    let numOcInter = objetoPedido.ocInter;

    if (objetoPedido.requiereInter === "Sí") {
        if (!numPedInter || numPedInter.trim() === "") numPedInter = "PENDIENTE";
        if (!numOcInter || numOcInter.trim() === "") numOcInter = "PENDIENTE";
    }

    const nuevaFila = cabeceras.map(cabecera => {
        switch (cabecera) {
            case 'FECHA RECEPCIÓN OC': return objetoPedido.fechaRegistro;
            case 'OC CLIENTE': return objetoPedido.ocCliente;
            case 'FECHA CREACIÓN PEDIDO': return objetoPedido.fechaRegistro;
            case 'PEDIDO CLIENTE': return objetoPedido.pedidoCliente;
            case 'PEDIDO INTER': return numPedInter;
            case 'OC INTER': return numOcInter;
            case 'ORG': return objetoPedido.org;
            case 'CLIENTE': return objetoPedido.cliente;
            case 'MATERIAL': return objetoPedido.material;
            case 'SKU PQ': return objetoPedido.productoSKUPQ;
            case 'SKU DJP': return objetoPedido.productoSKUDJP;
            case 'PRODUCTO': return objetoPedido.productoNombre;
            case 'ALIAS': return objetoPedido.alias;
            case 'CANTIDAD': return objetoPedido.cantidad;
            case 'UM': return objetoPedido.um;
            case 'ENTREGA PROGRAMADA': return objetoPedido.fechaProg;
            case 'FECHA DE ENTREGA': return objetoPedido.fechaEntregaReal || ""; // Agregado para Fletera
            case 'ORIGEN': return objetoPedido.origen;
            case 'TIPO DE ENVÍO':
            case 'TIPO DE ENVIO': return objetoPedido.tipoEnvio;
            case 'ESTATUS': return objetoPedido.estatus;
            case 'ESTATUS DE TIEMPO': return "";
            case 'PEDIDO INTER REQUERIDO': return objetoPedido.requiereInter;
            default: return "";
        }
    });
    sheet.appendRow(nuevaFila);
    return { success: true };
}

function actualizarPedidoDesdeDrawer(filaId, estatus, fechaEntrega) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MONITOREO DE PEDIDOS');
    const cabeceras = sheet.getDataRange().getValues()[0].map(c => String(c).trim().toUpperCase());

    const idxEstatus = cabeceras.indexOf('ESTATUS');
    const idxEntrega = cabeceras.indexOf('FECHA DE ENTREGA');

    if (idxEstatus === -1 || idxEntrega === -1) {
        throw new Error("No se encontraron las columnas 'ESTATUS' o 'FECHA DE ENTREGA'. Verifica los encabezados.");
    }

    // Actualizamos las celdas directamente usando el ID de fila
    sheet.getRange(filaId, idxEstatus + 1).setValue(estatus);
    sheet.getRange(filaId, idxEntrega + 1).setValue(fechaEntrega || "");

    return { success: true };
}

function toggleDrawerFecha() {
    const estatus = document.getElementById('drawerEstatus').value;
    const inputFecha = document.getElementById('drawerFechaEntrega');

    if (estatus === 'Completado') {
        inputFecha.disabled = false;
        // Si no tenía fecha, le sugerimos la fecha del día de hoy
        if (!inputFecha.value) inputFecha.value = new Date().toISOString().split('T')[0];
    } else {
        inputFecha.value = ""; // Limpia la fecha porque no se ha entregado
        inputFecha.disabled = true; // Bloquea el campo para evitar errores
    }
}

function eliminarPedido(filaId) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MONITOREO DE PEDIDOS');
    // Como le pasamos el filaId real, borra la fila exacta de la hoja
    sheet.deleteRow(filaId);
    return { success: true };
}

function actualizarPedido(filaId, objetoPedido) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MONITOREO DE PEDIDOS');
    const cabeceras = sheet.getDataRange().getValues()[0].map(c => String(c).trim().toUpperCase());
    const filaActual = sheet.getRange(filaId, 1, 1, cabeceras.length).getValues()[0];

    const filaActualizada = cabeceras.map((cabecera, index) => {
        switch (cabecera) {
            case 'ORG': return objetoPedido.org;
            case 'CLIENTE': return objetoPedido.cliente;
            case 'OC CLIENTE': return objetoPedido.ocCliente;
            case 'PEDIDO CLIENTE': return objetoPedido.pedidoCliente;
            case 'ENTREGA PROGRAMADA': return objetoPedido.fechaProg;
            case 'FECHA DE ENTREGA': return objetoPedido.fechaEntregaReal || filaActual[index];
            case 'ORIGEN': return objetoPedido.origen;
            case 'TIPO DE ENVÍO':
            case 'TIPO DE ENVIO': return objetoPedido.tipoEnvio;
            case 'ESTATUS': return objetoPedido.estatus;
            default: return filaActual[index];
        }
    });

    sheet.getRange(filaId, 1, 1, filaActualizada.length).setValues([filaActualizada]);
    return { success: true };
}