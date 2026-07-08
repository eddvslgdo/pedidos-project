function doGet() {
    return HtmlService.createTemplateFromFile('modal')
        .evaluate()
        .setTitle('Gestor Integral de Pedidos - Polaquimia')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Obtiene métricas avanzadas, registros detallados y catálogos para filtros
 */
function obtenerMetricasKPIs() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MONITOREO DE PEDIDOS');
    if (!sheet) return { error: "No se encontró la pestaña 'MONITOREO DE PEDIDOS'" };

    const datos = sheet.getDataRange().getValues();
    const cabeceras = datos[0];
    const filas = datos.slice(1);

    // Índices mapeados por cabecera exacta
    const idxOC = cabeceras.indexOf('OC CLIENTE');
    const idxPedido = cabeceras.indexOf('PEDIDO CLIENTE');
    const idxPedidoInter = cabeceras.indexOf('PEDIDO INTER');
    const idxCliente = cabeceras.indexOf('CLIENTE');
    const idxProducto = cabeceras.indexOf('PRODUCTO');
    const idxProg = cabeceras.indexOf('ENTREGA PROGRAMADA');
    const idxEntrega = cabeceras.indexOf('FECHA DE ENTREGA');
    const idxEnvio = cabeceras.indexOf('TIPO DE ENVÍO');
    const idxMotivo = cabeceras.indexOf('MOTIVO DE RETRASO');

    let ocsTotales = 0, pedidosConvertidos = 0;
    let aTiempo = 0, retraso = 0, enTransito = 0;

    const conteoClientes = {};
    const conteoProductos = {};
    const historialPedidos = [];
    const clientesUnicos = new Set();
    const productosUnicos = new Set();

    filas.forEach((fila, nFila) => {
        const oc = fila[idxOC];
        const pedido = fila[idxPedido];
        const pedInter = fila[idxPedidoInter];
        const cliente = fila[idxCliente];
        const producto = fila[idxProducto];
        const fProg = fila[idxProg];
        const fEntrega = fila[idxEntrega];
        const tEnvio = String(fila[idxEnvio]).toUpperCase();
        const motivo = fila[idxMotivo] || "No especificado";

        if (!oc && !pedido) return;

        if (oc) ocsTotales++;
        if (pedido) pedidosConvertidos++;
        if (cliente) { conteoClientes[cliente] = (conteoClientes[cliente] || 0) + 1; clientesUnicos.add(cliente); }
        if (producto) { conteoProductos[producto] = (conteoProductos[producto] || 0) + 1; productosUnicos.add(producto); }

        let estadoEntrega = "En Tránsito";

        if (pedido) {
            if (!fEntrega || fEntrega === "") {
                enTransito++;
            } else if (tEnvio === "FLETERA") {
                // Regla de negocio: Si es fletera, es exención de responsabilidad -> Cuenta como A Tiempo
                aTiempo++;
                estadoEntrega = "A Tiempo (Fletera)";
            } else {
                const dateProg = new Date(fProg);
                const dateEntrega = new Date(fEntrega);
                dateProg.setHours(0, 0, 0, 0);
                dateEntrega.setHours(0, 0, 0, 0);

                if (dateEntrega <= dateProg) {
                    aTiempo++;
                    estadoEntrega = "A Tiempo";
                } else {
                    retraso++;
                    estadoEntrega = "Atrasado";
                }
            }
        }

        historialPedidos.push({
            filaId: nFila + 2,
            oc,
            pedido,
            pedInter: pedInter || "PENDIENTE",
            cliente,
            producto,
            estadoEntrega,
            motivoRetraso: estadoEntrega === "Atrasado" ? motivo : "—"
        });
    });

    const topClientes = Object.entries(conteoClientes).sort((a, b) => b[1] - a[1]).slice(0, 3);
    const topProductos = Object.entries(conteoProductos).sort((a, b) => b[1] - a[1]).slice(0, 3);
    const totalConcluidos = aTiempo + retraso;

    return {
        ocsTotales,
        pedidosConvertidos,
        tasaConversionOC: ocsTotales > 0 ? ((pedidosConvertidos / ocsTotales) * 100).toFixed(1) : 0,
        aTiempo,
        retraso,
        enTransito,
        tasaEfectividad: totalConcluidos > 0 ? ((aTiempo / totalConcluidos) * 100).toFixed(1) : 0,
        topClientes,
        topProductos,
        historialPedidos,
        clientesLista: Array.from(clientesUnicos),
        productosLista: Array.from(productosUnicos)
    };
}

/**
 * Inserta un nuevo pedido validando flujos intercompañía
 */
function guardarNuevoPedido(objetoPedido) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MONITOREO DE PEDIDOS');
    if (!sheet) return { error: "Error al acceder a la hoja." };

    const cabeceras = sheet.getDataRange().getValues()[0];

    // Si requiere intercompañía pero no se capturó número, se registra explícitamente como "PENDIENTE INTER"
    let numPedInter = objetoPedido.pedidoInter;
    if (objetoPedido.requiereInter === "Sí" && (!numPedInter || numPedInter.trim() === "")) {
        numPedInter = "PENDIENTE INTER";
    }

    // Estructurar la nueva fila respetando el mapa de columnas
    const nuevaFila = cabeceras.map(cabecera => {
        switch (cabecera) {
            case 'FECHA RECEPCIÓN OC': return objetoPedido.fechaOc;
            case 'OC CLIENTE': return objetoPedido.ocCliente;
            case 'FECHA CREACIÓN PEDIDO': return objetoPedido.fechaPedido;
            case 'PEDIDO CLIENTE': return objetoPedido.pedidoCliente;
            case 'PEDIDO INTER': return numPedInter;
            case 'ORG': return objetoPedido.org;
            case 'CLIENTE': return objetoPedido.cliente;
            case 'PRODUCTO': return objetoPedido.producto;
            case 'CANTIDAD': return objetoPedido.cantidad;
            case 'UM': return objetoPedido.um;
            case 'ENTREGA PROGRAMADA': return objetoPedido.fechaProg;
            case 'ORIGEN': return objetoPedido.origen;
            case 'TIPO DE ENVÍO': return objetoPedido.tipoEnvio;
            case 'PEDIDO INTER REQUERIDO': return objetoPedido.requiereInter;
            default: return "";
        }
    });

    sheet.appendRow(nuevaFila);
    return { success: true };
}