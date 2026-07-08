function doGet() {
    return HtmlService.createTemplateFromFile('modal')
        .evaluate()
        .setTitle('Gestor Integral de Pedidos - Polaquimia')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obtenerMetricasKPIs() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MONITOREO DE PEDIDOS');
    if (!sheet) return { error: "No se encontró la pestaña 'MONITOREO DE PEDIDOS'" };

    const datos = sheet.getDataRange().getValues();
    const cabeceras = datos[0].map(c => String(c).trim().toUpperCase());
    const filas = datos.slice(1);

    const getIdx = (nombre) => cabeceras.indexOf(nombre);

    const idxOC = getIdx('OC CLIENTE');
    const idxPedido = getIdx('PEDIDO CLIENTE');
    const idxCreacion = getIdx('FECHA CREACIÓN PEDIDO');

    // CAMPOS INTERCOMPAÑIA
    const idxPedidoInter = getIdx('PEDIDO INTER');
    const idxOcInter = getIdx('OC INTER');

    const idxOrg = getIdx('ORG');
    const idxCliente = getIdx('CLIENTE');
    const idxMaterial = getIdx('MATERIAL');

    const idxSkuPq = getIdx('SKU PQ');
    const idxProducto = getIdx('PRODUCTO');

    const idxCantidad = getIdx('CANTIDAD');
    const idxUm = getIdx('UM');

    const idxProg = getIdx('ENTREGA PROGRAMADA');
    const idxEntrega = getIdx('FECHA DE ENTREGA');
    const idxEstatus = getIdx('ESTATUS');
    const idxEstatusTiempo = getIdx('ESTATUS DE TIEMPO');

    const idxOrigen = getIdx('ORIGEN');
    const idxEnvio = getIdx('TIPO DE ENVIO') !== -1 ? getIdx('TIPO DE ENVIO') : getIdx('TIPO DE ENVÍO');

    let ocsTotales = 0, pedidosConvertidos = 0;
    const historialPedidos = [];
    const clientesUnicos = new Set();
    const productosUnicos = new Set();
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

    filas.forEach((fila, nFila) => {
        const oc = idxOC !== -1 ? fila[idxOC] : "";
        const pedido = idxPedido !== -1 ? fila[idxPedido] : "";
        const pedInter = idxPedidoInter !== -1 ? fila[idxPedidoInter] : "";
        const ocInter = idxOcInter !== -1 ? fila[idxOcInter] : "";

        const org = idxOrg !== -1 ? fila[idxOrg] : "";
        const cliente = idxCliente !== -1 ? fila[idxCliente] : "";
        const material = idxMaterial !== -1 ? fila[idxMaterial] : "";

        const skuPq = idxSkuPq !== -1 ? fila[idxSkuPq] : "";
        const productoNombre = idxProducto !== -1 ? fila[idxProducto] : "";

        const cantidad = idxCantidad !== -1 ? (parseFloat(fila[idxCantidad]) || 0) : 0;
        const um = idxUm !== -1 ? fila[idxUm] : "";

        const estatus = idxEstatus !== -1 ? fila[idxEstatus] : "";
        const estatusTiempo = idxEstatusTiempo !== -1 ? fila[idxEstatusTiempo] : "";

        if (!oc && !pedido && !pedInter && !ocInter) return;

        if (oc || ocInter) ocsTotales++;
        if (pedido || pedInter) {
            pedidosConvertidos++;
            if (cliente) clientesUnicos.add(cliente);
            if (productoNombre) productosUnicos.add(productoNombre);
            if (org) orgsUnicas.add(org);
            if (material) materialesUnicos.add(material);
        }

        let estadoEntrega = "En Proceso";
        if (estatus === "Completado") {
            estadoEntrega = estatusTiempo ? estatusTiempo : "Completado";
        } else if (estatus === "En Proceso") {
            estadoEntrega = "En Proceso";
        }

        historialPedidos.push({
            filaId: nFila + 2,
            oc, pedido,
            pedInter: pedInter || "—",
            ocInter: ocInter || "—",
            cliente, org, material, cantidad, um,
            productoNombre: productoNombre,
            productoSKU: skuPq,
            estadoEntrega,
            origen: idxOrigen !== -1 ? fila[idxOrigen] : "",
            tipoEnvio: idxEnvio !== -1 ? fila[idxEnvio] : "",
            fechaCreacionISO: formatoISO(idxCreacion !== -1 ? fila[idxCreacion] : ""),
            fechaCreacion: formatearFechaStr(idxCreacion !== -1 ? fila[idxCreacion] : ""),
            fechaProg: formatearFechaStr(idxProg !== -1 ? fila[idxProg] : ""),
            fechaEntrega: formatearFechaStr(idxEntrega !== -1 ? fila[idxEntrega] : ""),
            fechaProgRaw: formatoISO(idxProg !== -1 ? fila[idxProg] : "")
        });
    });

    return {
        tasaConversionOC: ocsTotales > 0 ? ((pedidosConvertidos / ocsTotales) * 100).toFixed(1) : 0,
        historialPedidos,
        clientesLista: Array.from(clientesUnicos),
        productosLista: Array.from(productosUnicos),
        orgsLista: Array.from(orgsUnicas),
        materialesLista: Array.from(materialesUnicos)
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
            case 'SKU PQ': return objetoPedido.productoSKU;
            case 'PRODUCTO': return objetoPedido.productoNombre;
            case 'CANTIDAD': return objetoPedido.cantidad;
            case 'UM': return objetoPedido.um;
            case 'ENTREGA PROGRAMADA': return objetoPedido.fechaProg;
            case 'ORIGEN': return objetoPedido.origen;
            case 'TIPO DE ENVÍO':
            case 'TIPO DE ENVIO': return objetoPedido.tipoEnvio;
            case 'ESTATUS': return objetoPedido.estatus;
            case 'ESTATUS DE TIEMPO': return "";
            case 'PEDIDO INTER REQUERIDO': return objetoPedido.requiereInter; //Añadido para mantener consistencia
            default: return "";
        }
    });
    sheet.appendRow(nuevaFila);
    return { success: true };
}

function actualizarPedido(filaId, objetoPedido) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MONITOREO DE PEDIDOS');
    const cabeceras = sheet.getDataRange().getValues()[0].map(c => String(c).trim().toUpperCase());
    const filaActual = sheet.getRange(filaId, 1, 1, cabeceras.length).getValues()[0];

    const filaActualizada = cabeceras.map((cabecera, index) => {
        switch (cabecera) {
            case 'OC CLIENTE': return objetoPedido.ocCliente || filaActual[index];
            case 'PEDIDO CLIENTE': return objetoPedido.pedidoCliente || filaActual[index];
            case 'ORG': return objetoPedido.org || filaActual[index];
            case 'CLIENTE': return objetoPedido.cliente || filaActual[index];
            case 'ENTREGA PROGRAMADA': return objetoPedido.fechaProg || filaActual[index];
            case 'FECHA DE ENTREGA': return objetoPedido.fechaEntregaReal || filaActual[index];
            case 'ORIGEN': return objetoPedido.origen || filaActual[index];
            case 'TIPO DE ENVÍO':
            case 'TIPO DE ENVIO': return objetoPedido.tipoEnvio || filaActual[index];
            case 'ESTATUS': return objetoPedido.estatus || filaActual[index];
            default: return filaActual[index];
        }
    });

    sheet.getRange(filaId, 1, 1, filaActualizada.length).setValues([filaActualizada]);
    return { success: true };
}