//Funcion para crear un indice con vinculo a todas mis pestañas
function indiceDinamico() {
  //Configuracion
  const nombreInicio = "Inicio";
  var misHojas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var hojaInicio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreInicio);
  var i=2;
  
  //Borrar todo el contenido antes de actualizar el indice
  hojaInicio.getRange(2,1,hojaInicio.getLastRow()).clearContent();
  
  //Ciclo que recorre todas las pestañas actuales
  misHojas.forEach(function(hoja){
    
    //Condicional para que no tome la pestaña "Inicio"
    if(hoja.getName()!=nombreInicio){
      
      //Agregue la formula HIPERVINCULO con el nombre y la Id de cada pestaña
      var formula = '=HYPERLINK("#gid='+hoja.getSheetId()+'";"'+hoja.getName()+'")'
      hojaInicio.getRange(i,1).setFormula(formula)
      
      //Agregue la formula HIPERVINCULO con la pestaña Inicio a cada pestaña, para poder devolverse
      hoja.getRange(1,1).setFormula('=HYPERLINK("#gid=0";"Ir a Inicio")')
      i++;
    } // Cierre If Inicio
  }) // Cierre Ciclo
}
