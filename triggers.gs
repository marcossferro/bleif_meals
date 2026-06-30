function onOpen(){
  var menu = SpreadsheetApp.getUi().createMenu('Scripts')
  .addItem('Crear Archivo Mes Siguiente','duplicarWorkbookYLimpiar')
  
  menu.addToUi();
}
