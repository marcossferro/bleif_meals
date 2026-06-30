const SHEETID = SpreadsheetApp.getActiveSpreadsheet().getId();
const recetaTemplate = SpreadsheetApp.openById(SHEETID).getSheetByName("Recetas Template");
const listadoRecetas = SpreadsheetApp.openById(SHEETID).getSheetByName("Recetas ---->");


function duplicarPestanaConNombre() {
  var ss = SpreadsheetApp.openById(SHEETID);
  SpreadsheetApp.getActiveSpreadsheet(); // Obtiene la hoja activa

  // Mostrar un prompt para ingresar el nuevo nombre
  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.prompt(
    "Duplicar pestaña", 
    "Ingrese el nombre de la nueva pestaña:", 
    ui.ButtonSet.OK_CANCEL);

  // Si el usuario cancela o deja el campo vacío, salir del script
  if (respuesta.getSelectedButton() === ui.Button.CANCEL || respuesta.getResponseText().trim() === "") {
    ui.alert("Operación cancelada.");
    return;
  }

  var nuevoNombre = respuesta.getResponseText().trim();

  // Verificar si el nombre ya existe y evitar duplicados
  if(ss.getSheetByName(nuevoNombre) != null){
    ui.alert('La receta que queres crear ya Existe')
    return;
  }

  // Duplicar la hoja y asignarle el nuevo nombre
  var nuevaHoja = recetaTemplate.copyTo(ss);
  nuevaHoja.setName(nuevoNombre);

  // Confirmación
  ui.alert("Creaste la receta '" + nuevoNombre + "'!!");

  var rangosProtegidos = ["A:A", "C:C", "E:E"]; 
  
  // Seleccionar los rangos
  var rangoList = nuevaHoja.getRangeList(rangosProtegidos).getRanges(); // Obtener los rangos como array

  rangoList.forEach(rango => {
    var proteccion = rango.protect(); // Crear protección
    proteccion.setDescription("No Tocar"); // Descripción opcional
    proteccion.setWarningOnly(true); // Muestra advertencia en vez de bloquear completamente
  });

  nuevaHoja.getRange('B3:B').clearContent();
  nuevaHoja.getRange('D3:D').clearContent();
  
  SpreadsheetApp.flush();

  var hojas = ss.getSheets();
  var excluir = ["AUXILIAR", "Codificacion ", "PRODUCCION Y VENTA X COD", "Produccion", "Ventas", "Stock y Compras", "Recetas", "Modo de Uso", "Recetas Template", "Recetas ---->"]; // Nombres de pestañas a ignorar


  var nombres = hojas
    .map(hoja => hoja.getName()) // Obtener nombres de todas las pestañas
    .filter(nombre => !excluir.includes(nombre)) // Filtrar las pestañas a excluir
    .sort(); // Ordenar alfabéticamente
   
  if(nombres.length > 0) {
    listadoRecetas.getRange(2, 1, nombres.length, 1).setValues(nombres.map(nombre => [nombre])); // Pegarlos en la columna A
  } else {
    listadoRecetas.getRange(2, 1).setValue("No hay pestañas disponibles.");
  }
}

function ordenarPestanasPorNombre() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const pivotIndex = sheets.findIndex(s => s.getName() === 'Recetas ---->');
  if (pivotIndex === -1) {
    SpreadsheetApp.getUi().alert('No se encontró la pestaña "Recetas ---->"');
    return;
  }

  const toSort = sheets.slice(pivotIndex + 1);
  toSort.sort((a, b) => a.getName().localeCompare(b.getName(), 'es', { sensitivity: 'base' }));

  toSort.forEach((sheet, i) => {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(pivotIndex + 2 + i); // 1-based, +2 para dejar pivot en su lugar
  });
}
