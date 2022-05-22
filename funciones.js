//variable donde guardaremos todos los estilos

var estilos_sheet = PropertiesService.getDocumentProperties();

function onOpen() {
  
  SpreadsheetApp.getUi().createMenu("Aulaenlanube")
  .addItem("Mostrar barra Lateral","mostrarBarraLateral")
  .addToUi();

}
// crea la barra


function mostrarBarraLateral()
{
  var barra = HtmlService.createHtmlOutputFromFile("BarraLateral").setTitle("Barra lateral Aulaenlanube");
  SpreadsheetApp.getUi().showSidebar(barra);
}
//ejecuta la barra y la conecta con html pal diseño

function aplicarEstilo10(){
  // hojaactual sea la hoja de la hoja de calculo activa
  var hojaActual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  var celdas = hojaActual.getActiveRange();

  celdas.setBackground("blue")  //color de celda
        .setFontColor("white") //color de letra
        .setHorizontalAlignment("center")
        .setFontWeight("bold")
        .setValue("Estilo1");

}

//la hoja actual seria la clase donde se consigue la hoja de excel activa, con la hoja activa y el rango de celdas activa

function aplicarEstilo(estilo){
  // hojaactual sea la hoja de la hoja de calculo activa
  var hojaActual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  var celdas = hojaActual.getActiveRange();

  celdas.setFontColor(estilos_sheet.getProperty("colorLetra"+estilo))   //color de letra
        .setBackground(estilos_sheet.getProperty("colorFondo"+estilo))  //color de celda
        .setFontSize(estilos_sheet.getProperty("size"+estilo)); 

}
// aplica los estilos conseguidos en guardar estilo

//Lo mismo que arriba pero conseguir las propiedades 

function guardarEstilo(estilo)
{
  // hojaactual sea la hoja de la hoja de calculo activa
  var celda = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();

  estilos_sheet.setProperty("colorLetra"+estilo, celda.getFontColor())
               .setProperty("colorFondo"+estilo, celda.getBackground())
               .setProperty("size"+estilo, celda.getFontSize()+"");

  return {  colorFondo: estilos_sheet.getProperty("colorFondo"+estilo),
            colorLetra: estilos_sheet.getProperty("colorLetra"+estilo)};
}

// setproperty agarra key primero, y el segundo es para conseguir aquellos valores de la hoja activa, tonces colorletra es variable
// estilos_sheet es el objeto con esas caracteristicas y se devuelven dsp a aplicarestilo
// en el return, estos se devuelven a la pag html y colorFondo del return consigue el atributo del objeto estilos_sheet, devuelve un 
// diccionario


function cargarEstilos(){

  return estilos_sheet.getProperties();
}


function eliminarEstilo(estilo)
{
  estilos_sheet.deleteProperty("colorLetra"+estilo);
  estilos_sheet.deleteProperty("colorFondo"+estilo);
  estilos_sheet.deleteProperty("size"+estilo);

}

function borrarEstilo()
{
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear({formatOnly: true});
}


function borrarTodo()
{
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear();
}










