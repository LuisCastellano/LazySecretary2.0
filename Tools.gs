//////////////////////////////////////////////////////////////////////////////////////
//Capturar unidades organizativas: Las ordena y devuelve el objeto con todas las U.O
//////////////////////////////////////////////////////////////////////////////////////
function getUO() {
    try {
        var orgUnits = AdminDirectory.Orgunits.list("my_customer", {
            type: "all"
        }).organizationUnits

        orgUnits.sort(orderUO);
        var obj_orgUnits = Object.keys(orgUnits).map(function (k) {
            return orgUnits[k]
        });
        //Logger.log( obj_orgUnits[1].name); //Se recorre el objeto y se extraen los datos.
        /*{ --Recursos orgunits--
         "kind": "admin#directory#orgUnit",
         "etag": etag,
          "name": string,
          "description": string,
          "orgUnitPath": string,
          "orgUnitId": string,
          "parentOrgUnitPath": string,
          "parentOrgUnitId": string,
          "blockInheritance": boolean
         }*/
        return obj_orgUnits;
    } catch (e) {
        if (e.message == 'Cannot call method "sort" of undefined.') {
            //aviso(translate("No se encontró ninguna U.O."));
        }
        saveError(e)
        return null;
    }
}
//Ordena las UO.
function orderUO(a, b) {
    if (a.orgUnitPath.toLowerCase() < b.orgUnitPath.toLowerCase()) {
        return -1;
    }
    if (a.orgUnitPath.toLowerCase() > b.orgUnitPath.toLowerCase()) {
        return 1;
    }
    return 0;
}








///////////////////////////////////////////////////////////////////////////////////////////
//Crea las u.o: recibe como parametro un array objetos.
///////////////////////////////////////////////////////////////////////////////////////////
function createUO(resource) {
    try {
        if (resource.length == 0.0 || resource == undefined || resource == null) { //Si el array de recursos es 0, indefinido o nulo, informa al usuario
            Logger.log(translate("Información de U.O inválidas o no encontrada")); //MOSTRAR AL USUARIO O REGISTRAR.
            return;
        }

        for (var i = 0; i < resource.length; i++) {
            try {

                var newUO = JSON.parse(resource[i]);
                AdminDirectory.Orgunits.insert(newUO, "my_customer"); //Crea la u.o

            } catch (e) {

                //Tratamiento de errores:
                switch (e.message) {
                    case "Invalid Ou Id":
                        Logger.log(translate("En la ruta '") + newUO.parentOrgUnitPath + translate("' ya existe una U.O con el nombre '") + newUO.name + "'."); //MOSTRAR AL USUARIO O REGISTRAR.
                        continue;
                        break;

                    case "Invalid Parent Orgunit Id":
                        Logger.log(translate("La ruta '") + newUO.parentOrgUnitPath + translate("' no existe.")); //MOSTRAR AL USUARIO O REGISTRAR.
                        continue;
                        break;

                    default:
                        Logger.log(translate("Error inesperado:\n No se pudo crear la U.O '" + newUO.name + "'.")); //MOSTRAR AL USUARIO O REGISTRAR.
                        saveError(e);
                }
            } // end catch
        } // end for

    } catch (e) {
        saveError(e);
    }
}



///////////////////////////////////////////////////////
// Devuelve la cuota diaria de mails que queda
///////////////////////////////////////////////////////
function checkDailyQuota() {
    return MailApp.getRemainingDailyQuota();
}







///////////////////////////////////////////////////////////////////
//Crea la hoja principal del complemento.
//////////////////////////////////////////////////////////////////
function createFirstSheet(){
 var name_workSheet = (translate('Usuarios')+'_LazySecretary'),
     ss = SpreadsheetApp.getActiveSpreadsheet();
  if(ss.getSheetByName(name_workSheet)){return;}// si existe no se crea
  set_userProperties(name_workSheet, 0, false, true);
 var colorRGB = [0.0, 0.5, 1.0],
     firstSheet = getOrCreateSheet(name_workSheet, colorRGB);
  generate_header(firstSheet);
  return;
}








////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Crea las hojas de Errores y usuarios creados.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function prepare_create_and_errors() {
  try{
     //**TRADUCCIONES**//
    var arr_string = ["Usuarios creados", "Errores detectados"];
    arr_string = translateArray(arr_string);
    ///////////////////////////
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if(ss.getSheetByName(arr_string[0]) && ss.getSheetByName(arr_string[1])){
      return;
    }
    // si existe no se crea
    if(!ss.getSheetByName(arr_string[0])){
      var createdSheet = getOrCreateSheet(arr_string[0], [0.01,0.75,0.24]);// crea la hoja de usuarios creados si no existe.
    }
    
    // si existe no se crea
    if(!ss.getSheetByName(arr_string[1])){
      var errorSheet = getOrCreateSheet(arr_string[1], [0.9,0.17,0.31]);// crea la hoja de errores detectados si no existe.    
    }
   
      
    generate_header(createdSheet);
    generate_header(errorSheet);
    SpreadsheetApp.flush();
    return "succesful";
  }catch(e){
    saveError(e);
    aviso("ERROR prepareWorkSheet: "+e);
    return "failed"
  }
}



//////////////////////////////////////////////////////////
// configura la spreadsheet con las hojas iniciales.
//////////////////////////////////////////////////////////
function initialize() {
  try{
  createFirstSheet();
  prepare_create_and_errors();
  return true;
  }catch(e){
    saveError(e);
  }
  
}






///////////////////////////////////////////////////////////////////////////////////////////
// Devuelve los alias de dominio
///////////////////////////////////////////////////////////////////////////////////////////
function get_domain_aliases() {
    var v_domain_aliases = AdminDirectory.DomainAliases.list("my_customer").domainAliases;
    return v_domain_aliases;
}

///////////////////////////////////////////////////////////////////////////////////////////
// Devuelve el nombre de dominio principal
///////////////////////////////////////////////////////////////////////////////////////////
function get_domain() {
    var domain = AdminDirectory.Domains.list("my_customer").domains;
    return (domain[0]);
}





///////////////////////////////////////////////////////////////////////////////////////////
// Genera una cabecera con los diferentes valores a rellenar en la página principal
///////////////////////////////////////////////////////////////////////////////////////////
function generate_header(workSheet) {
    try {
     
      //TRADUCCIONES
        var header_values = ['Nombre', 'Apellidos', 'Contraseña', 'Dominio', 'Nuevo email', 'Email de contacto', 'U.O destino', 'Estado', 'Administrador', 'Fecha de creación'],
            createdSheet = translate("Usuarios creados");
            header_values = translateArray(header_values);      
            header_values.push('V2P', 'Id Googe');
      ////////////////////////////////////////////////
      
      //Añade los valores de la cabecera a las celdas
        for (var i = 0; i < header_values.length; i++) {
             workSheet.getRange(1,i+1).setValue(header_values[i]);
        }
      //Obtiene el numero correspondiente a la columna.
      var colEstado = getColIndexByName(header_values[7], workSheet),
          colIdGoogle = getColIndexByName(header_values[11], workSheet);
      
      //Cambia el ancho de las columnas
       for ( i = 0; i < header_values.length; i++) {
          // Si es la columna de estado, el ancho es mayor.
      if((i + 1) == colEstado){
            workSheet.setColumnWidth(i + 1,400); //Modifica el tamaño de la columna de estado. 
          }else{
            workSheet.setColumnWidth(i + 1,200); //Modifica el tamaño de la columna. 
          }
       }
      
      //FORMATO
     // Elimina el DV, modifica la fuente, alinea el contenido de la columna y modifica el color.
        workSheet.getRange(1,1,1,workSheet.getLastColumn()).clearDataValidations().setFontSize(15).setHorizontalAlignment('center').setWrap(true).setFontColor('#fff');
        workSheet.getRange(1,1,1,(colEstado-1)).setBackground('#03a9f4');//blue
        workSheet.getRange(1,colEstado,1,1).setBackground('#ef5350');//red
        workSheet.getRange(1,(colEstado+1),1,4).setBackground('#66bb6a');//green
      // Añade una nota en la columna V2P.
        workSheet.getRange(1,getColIndexByName(header_values[10], workSheet)).setNote(translate("Verificación en dos pasos"));
        
      //////////////////////
        workSheet.setFrozenRows(1);
      if(workSheet.getName() != createdSheet){
        build_rule_dataV_domain(workSheet);
        build_rule_dataV_uo(workSheet);
      }
        SpreadsheetApp.flush();
        return;
        
    } catch (e) {
        saveError(e);
      return;
    }

}





///////////////////////////////////////////////////////////////////////////////////////////
// Construye las reglas del dataValidation de dominios y lo agrega al rango.
///////////////////////////////////////////////////////////////////////////////////////////
function build_rule_dataV_domain(sheet) {
    var colDomain = getColIndexByName(translate('Dominio'), sheet),
        range_DV = sheet.getRange(2, colDomain, sheet.getMaxRows()),
        alias_domains = get_domain_aliases(),
        domain = get_domain(),
        v_domain = [];
    //Agrega el dominio principal.
    v_domain.push(domain.domainName);
    //Agrega los alias de dominio.
    for (var i = 0; i < alias_domains.length; i++) {
        v_domain.push(alias_domains[i].domainAliasName);
    }
    //Crea la regla del dataValidation.
    var rule_action = SpreadsheetApp.newDataValidation().requireValueInList(v_domain).build();
    range_DV.setDataValidation(rule_action); //Agrega el DataValidation al rango.
    return;
}




///////////////////////////////////////////////////////////////////////////////////////////
// Construye las reglas del dataValidation de U.O y lo agrega al rango.
///////////////////////////////////////////////////////////////////////////////////////////
function build_rule_dataV_uo(sheet) {
    var obj_orgUnits = getUO(),
        colUO = getColIndexByName(translate('U.O destino'), sheet),
        v_uo = [];
  
   if(obj_orgUnits == null){
      v_uo.push('/');
   }else{
    //rellena el array con las rutas de la uo.
    for (var i = 0; i < obj_orgUnits.length; i++) {
       
        v_uo.push(obj_orgUnits[i].orgUnitPath);
    }
   }
    var range_DV = sheet.getRange(2, colUO, sheet.getMaxRows()),
        //Crea la regla del dataValidation.
        rule_action = SpreadsheetApp.newDataValidation().requireValueInList(v_uo).build();
    range_DV.setDataValidation(rule_action); //Agrega el DataValidation al rango.

}


/////////////////////////////////////////////////////////////////////////////////////////
// Comprueba si existe la hoja, si existe, la devuelve, si no, la crea y la devuelve.
/////////////////////////////////////////////////////////////////////////////////////////
function getOrCreateSheet(name_sheet, tabColor) {
  try{
    var ss = SpreadsheetApp.getActiveSpreadsheet();
        
    if (ss.getSheetByName(name_sheet)) {// si existe la hoja, la devuelve.
        return ss.getSheetByName(name_sheet);
    }
    // Crea la hoja si no existe.
    var newSheet = addSheet(name_sheet, tabColor);
    return newSheet;
  }catch(e){
    aviso("error OrCreate: "+e.message+e.fileName+e.lineNumber);
    saveError(e);
    return;
  }
}


/////////////////////////////////////////////////////////////////////////////////////////
// Copia la cabecera de la hoja activa a la hoja de destino dada.
/////////////////////////////////////////////////////////////////////////////////////////
function copyRange(sheet, rangeToCopy, targetSheet) {
    rangeToCopy.copyTo(targetSheet.getRange(targetSheet.getLastRow() + 1, 1));
    return;
}



/////////////////////////////////////////////////////////////////////////////////////////
// Actualiza los data validations de la hoja de trabajo principal.
/////////////////////////////////////////////////////////////////////////////////////////
function refresh_DV(worksheet_name){
var worksheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(worksheet_name);
  build_rule_dataV_uo(worksheet);
  build_rule_dataV_domain(worksheet);
    }

