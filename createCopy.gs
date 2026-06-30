// CONFIGURACIÓN — ajustá esto según tus hojas y rangos
// Cada entrada: nombre exacto de la pestaña -> array de rangos A1 a borrar
const RANGOS_A_BORRAR = {
  'Ventas': ['A3:B', 'E3:G', 'L3:BY'],
  'Produccion': ['F8:F12', 'L8:BW12'],
  'Stock y Compras': ['G4:M'],
  // 'NombreDeLaHoja': ['A1:D10', 'G1:G20'],
};

// CONFIGURACIÓN — traspaso de valores (final del original -> inicial de la copia)
// "origen" = rango en el archivo ORIGINAL de donde se copian los valores.
// "destino" = rango en la COPIA donde se pegan esos valores (solo valores, sin fórmulas).
// Soporta columnas abiertas (ej: "M4:M") — calcula automáticamente la última fila con datos usando la columna B como referencia.
const TRASPASO_VALORES = [
  { hoja: 'Produccion', origen: 'H15:BW15', destino: 'H7:BW7' },
  { hoja: 'Stock y Compras', origen: 'M4:M', destino: 'G4:G' },
];

// FUNCIÓN PRINCIPAL
function duplicarWorkbookYLimpiar() {
  const ui = SpreadsheetApp.getUi();

  // 1. Obtener archivo original y calcular el próximo mes a partir de su nombre
  //    y armar el nombre del archivo automáticamente
  const archivoOriginalId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const archivoOriginal = DriveApp.getFileById(archivoOriginalId);
  const nombreArchivoOriginal = archivoOriginal.getName();

  const info = obtenerProximoMesInfo(nombreArchivoOriginal);
  const nuevoNombre = 'Gastos Bleif Meals - ' + info.nombreMes.toUpperCase() + ' ' + info.anio;

  // Confirmación antes de crear el archivo
  const confirmacion = ui.alert(
    'Confirmar creación',
    'Archivo original: "' + nombreArchivoOriginal + '"\n' +
    'Se va a crear el archivo:\n"' + nuevoNombre + '"\n\n¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirmacion !== ui.Button.YES) return;

  // 2. Ubicar el archivo original dentro de la estructura: Raíz -> Año -> Mes -> archivo
  //    Por lo tanto subimos dos niveles desde el archivo para llegar a la Raíz.
  const carpetaMesActual = obtenerPrimerPadre(archivoOriginal);     // ej: "07 - Julio"
  const carpetaAnioActual = obtenerPrimerPadre(carpetaMesActual);   // ej: "2026"
  const carpetaRaiz = obtenerPrimerPadre(carpetaAnioActual);        // contiene todas las carpetas de año

  // 3. Buscar o crear la carpeta Año/Mes correspondiente al PRÓXIMO mes
  const carpetaDestino = obtenerOCrearCarpetaDelMes(carpetaRaiz, info);

  // 4. Duplicar el archivo completo y moverlo a la carpeta del mes
  const copia = archivoOriginal.makeCopy(nuevoNombre, carpetaDestino);

  // 5. Abrir ambos spreadsheets
  const ssCopia = SpreadsheetApp.openById(copia.getId());
  const ssOriginal = SpreadsheetApp.openById(archivoOriginalId);

  // 6. Limpiar los rangos viejos en la copia
  limpiarRangos(ssCopia);

  // 7. Traspasar valores "final del original" -> "inicial de la copia"
  traspasarValores(ssOriginal, ssCopia);

  ui.alert('Listo. Copia creada: "' + nuevoNombre + '" en la carpeta "' + carpetaDestino.getName() + '".\nLink: ' + copia.getUrl());
}

// Calcula el "próximo mes" a partir del nombre del archivo ORIGINAL
// (ej: "Gastos Bleif Meals - JULIO 2026"), buscando el mes dentro del
// array "meses" y devolviendo el siguiente. Si el mes encontrado es
// Diciembre, el próximo es Enero del año siguiente.
// Devuelve { anio, mesIndex (0-11), nombreCarpetaMes, nombreMes }.
function obtenerProximoMesInfo(nombreArchivoOriginal) {
  const meses = [
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ];

  // Extraer año del nombre del archivo (ej: "JULIO 2026" -> 2026)
  const matchAnio = nombreArchivoOriginal.match(/\d{4}/);
  if (!matchAnio) {
    throw new Error('No se pudo encontrar el año en el nombre del archivo: "' + nombreArchivoOriginal + '"');
  }
  const anioActual = parseInt(matchAnio[0], 10);

  // Buscar el mes actual dentro del array, comparando en mayúsculas
  const nombreArchivoNormalizado = nombreArchivoOriginal.toUpperCase();
  const mesIndexActual = meses.findIndex(mes => nombreArchivoNormalizado.includes(mes.toUpperCase()));
  if (mesIndexActual === -1) {
    throw new Error('No se pudo encontrar el nombre del mes en el archivo: "' + nombreArchivoOriginal + '"');
  }

  // Calcular el próximo mes dentro del array (con cambio de año si corresponde)
  let anio = anioActual;
  let mesIndex = mesIndexActual + 1;
  if (mesIndex > 11) {
    mesIndex = 0;
    anio += 1;
  }

  // Formato "07 - Julio" (número de mes 1-12, con cero a la izquierda)
  const numeroMes = String(mesIndex + 1).padStart(2, '0');
  const nombreCarpetaMes = numeroMes + ' - ' + meses[mesIndex];

  return {
    anio: anio,
    mesIndex: mesIndex,
    nombreCarpetaMes: nombreCarpetaMes,
    nombreMes: meses[mesIndex]
  };
}

// Busca la carpeta Año -> Mes ("2026" -> "07 - Julio") dentro de la carpeta raíz dada, usando la info ya calculada del próximo mes. 
// Crea cualquier carpeta que falte en el camino.
function obtenerOCrearCarpetaDelMes(carpetaRaiz, info) {
  const nombreAnio = String(info.anio);

  // 1. Carpeta del año
  let carpetaAnio;
  const carpetasAnioExistentes = carpetaRaiz.getFoldersByName(nombreAnio);
  if (carpetasAnioExistentes.hasNext()) {
    carpetaAnio = carpetasAnioExistentes.next();
  } else {
    carpetaAnio = carpetaRaiz.createFolder(nombreAnio);
  }

  // 2. Carpeta del mes dentro del año
  let carpetaMes;
  const carpetasMesExistentes = carpetaAnio.getFoldersByName(info.nombreCarpetaMes);
  if (carpetasMesExistentes.hasNext()) {
    carpetaMes = carpetasMesExistentes.next();
  } else {
    carpetaMes = carpetaAnio.createFolder(info.nombreCarpetaMes);
  }

  return carpetaMes;
}

// Devuelve la primera carpeta padre de un archivo o carpeta de Drive.
// Si no tiene padre (está en la raíz de Drive), devuelve la raíz.
function obtenerPrimerPadre(itemDrive) {
  const padres = itemDrive.getParents();
  return padres.hasNext() ? padres.next() : DriveApp.getRootFolder();
}

// Recorre el objeto RANGOS_A_BORRAR y limpia cada rango en su hoja correspondiente.
// Si querés que SOLO se borren los valores (mantener formato/validaciones):
// true = borra solo contenido | false = borra contenido + formato
const SOLO_CONTENIDO = true;

function limpiarRangos(spreadsheet) {
  Object.keys(RANGOS_A_BORRAR).forEach(nombreHoja => {
    const hoja = spreadsheet.getSheetByName(nombreHoja);
    if (!hoja) {
      console.warn('No se encontró la hoja "' + nombreHoja + '" en la copia. Se omite.');
      return;
    }
    const rangos = RANGOS_A_BORRAR[nombreHoja];
    rangos.forEach(rangoA1 => {
      const rango = hoja.getRange(rangoA1);
      if (SOLO_CONTENIDO) {
        rango.clearContent();
      } else {
        rango.clear(); // contenido + formato
      }
    });
  });
}

// Recorre TRASPASO_VALORES: para cada entrada, lee los valores del rango
// "origen" en el spreadsheet ORIGINAL y los pega (solo valores) en el rango
// "destino" del spreadsheet COPIA.
function traspasarValores(ssOriginal, ssCopia) {
  TRASPASO_VALORES.forEach(item => {
    const hojaOrigen = ssOriginal.getSheetByName(item.hoja);
    const hojaDestino = ssCopia.getSheetByName(item.hoja);

    if (!hojaOrigen || !hojaDestino) {
      console.warn('No se encontró la hoja "' + item.hoja + '" en origen o destino. Se omite el traspaso.');
      return;
    }

    const rangoOrigenResuelto = resolverRangoAbierto(hojaOrigen, item.origen);
    const valores = hojaOrigen.getRange(rangoOrigenResuelto).getValues();

    const rangoDestinoResuelto = ajustarRangoDestinoADimensiones(item.destino, valores);
    hojaDestino.getRange(rangoDestinoResuelto).setValues(valores);
  });
}

// Si el rango tiene una columna abierta (ej: "M4:M", sin fila final),
// lo resuelve usando la última fila con contenido en la columna B de la hoja
// (columna de referencia confiable para determinar el largo de la tabla).
// Si el rango ya está cerrado (ej: "H15:BW15"), lo devuelve tal cual.
function resolverRangoAbierto(hoja, rangoA1) {
  const partes = rangoA1.split(':');
  if (partes.length !== 2) return rangoA1;

  const inicio = partes[0];
  const fin = partes[1];

  // Si "fin" no tiene número de fila (ej: "M"), es una columna abierta
  const finTieneFila = /\d/.test(fin);
  if (finTieneFila) return rangoA1;

  const colInicioMatch = inicio.match(/[A-Z]+/)[0];
  const filaInicio = parseInt(inicio.match(/\d+/)[0], 10);
  const colFin = fin.match(/[A-Z]+/)[0];

  // Última fila con datos, usando la columna B como referencia
  const ultimaFila = hoja.getRange('B1:B' + hoja.getMaxRows()).getValues()
    .filter(fila => fila[0] !== '' && fila[0] !== null)
    .length;

  const filaFinal = Math.max(ultimaFila, filaInicio); // por si no hay datos en B

  return colInicioMatch + filaInicio + ':' + colFin + filaFinal;
}

// Ajusta el rango destino para que tenga exactamente las mismas dimensiones
// (filas x columnas) que la matriz de valores a pegar, partiendo de la celda de inicio indicada en "destinoA1".
function ajustarRangoDestinoADimensiones(destinoA1, valores) {
  const numFilas = valores.length;
  const numCols = valores[0] ? valores[0].length : 1;

  const inicio = destinoA1.split(':')[0];
  const colInicio = inicio.match(/[A-Z]+/)[0];
  const filaInicio = parseInt(inicio.match(/\d+/)[0], 10);

  const filaFin = filaInicio + numFilas - 1;
  const colFin = columnaPorIndice(columnaAIndice(colInicio) + numCols - 1);

  return colInicio + filaInicio + ':' + colFin + filaFin;
}

// Convierte letra(s) de columna (ej: "BW") a índice numérico (1-based).
function columnaAIndice(col) {
  let indice = 0;
  for (let i = 0; i < col.length; i++) {
    indice = indice * 26 + (col.charCodeAt(i) - 64);
  }
  return indice;
}

// Convierte índice numérico (1-based) a letra(s) de columna (ej: 75 -> "BW").
function columnaPorIndice(indice) {
  let col = '';
  while (indice > 0) {
    const resto = (indice - 1) % 26;
    col = String.fromCharCode(65 + resto) + col;
    indice = Math.floor((indice - 1) / 26);
  }
  return col;
}
