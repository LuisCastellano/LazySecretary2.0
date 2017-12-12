///////////////////////////////////////////////////////////////////////////////////////////
// Comprueba si una direccion de correo está bien formada -- return true/false
///////////////////////////////////////////////////////////////////////////////////////////
function checkEmail(email) {
    var regEx = /^([\w-]+(?:\.[\w-]+)*)@((?:[\w-]+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)$/i;
    return (regEx.test(email));
}




////////////////////////////////////////////////////////////////////////////////////
// Recoge los datos de la hoja de calculo, a partir de la segunda fila
////////////////////////////////////////////////////////////////////////////////////
function getData(sheet) {
    try {
        var startRow = 2,
            numRows = (sheet.getLastRow() - 1),
            data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()),
            values = data.getValues();
        return values;

    } catch (e) {
        return null;
    }
}



////////////////////////////////////////////////////////////////
// Captura el lenguaje de la cuenta del usuario y lo devuelve.
////////////////////////////////////////////////////////////////
function getLang() {
    try {
        var lang = Session.getActiveUserLocale();
        return lang;

    } catch (e) {
        saveError(e);
        return 'es';
    }
}




////////////////////////////////////////////////////////////////////////////////////
// Traduce cadenas de texto para compatibilizar el complemento con otros idiomas
///////////////////////////////////////////////////////////////////////////////////
function translate(text_to_translate) {
    try {
        var idioma = getLang(); // Idioma del usuario

        if (idioma == 'es') {
            return text_to_translate;
        } else { //Si el idioma del usuario es español no se traduce.

            text_to_translate = LanguageApp.translate(text_to_translate, 'es', idioma);
            return text_to_translate;
        }
    } catch (e) {
        aviso("Translation failure \n Please restart the add-on.");
        saveError(e);
        return text_to_translate;
    }
}


////////////////////////////////////////////////////////////////////////////////////
// Traduce arrays de texto para compatibilizar el complemento con otros idiomas
///////////////////////////////////////////////////////////////////////////////////
function translateArray(array_to_translate) {
    try {
        var idioma = getLang(); // Idioma del usuario
        if (idioma == 'es') {
            return array_to_translate;
        } else { //Si el idioma del usuario es español no se traduce.
            for (var i = 0; i < array_to_translate.length; i++) {
                array_to_translate[i] = LanguageApp.translate(array_to_translate[i], 'es', idioma);
            }
            return array_to_translate;
        }
    } catch (e) {
        aviso("Translation failure \n Please restart the add-on.");
        saveError(e);
        return array_to_translate;
    }
}






////////////////////////////////////////////////////////////////////////////////////
///Muestra una ventana de alerta con el texto recibido por parametros.                        
///////////////////////////////////////////////////////////////////////////////////
function aviso(texto) {
    //Se ha agotado el tiempo mientras se esperaba la respuesta del usuario
    var strings = translateArray(['AVISO', texto]);// traduce el texto y el título.
    
    Browser.msgBox(strings[0], strings[1],Browser.Buttons.OK);
}





///////////////////////////////////////////////////////////////////////////////////////////////
// Muestra un mensaje de confirmacion preguntando el mensaje que se le pasa por parametros
//////////////////////////////////////////////////////////////////////////////////////////////
function askYesNo(mensaje) {
    var ui = SpreadsheetApp.getUi(),
        result = ui.alert(
            translate("Por favor, confirma"),
            translate(mensaje),
            ui.ButtonSet.YES_NO_CANCEL);

    if (result == ui.Button.YES) {
        return true;
    } else if (result == ui.Button.NO) {
        return false;
    }else{
      return "cancel";
    }
}


/////////////////////////////////////////////////////////////
// Comprueba si todos los campos estan rellenados.
////////////////////////////////////////////////////////////
function checkData(new_email, nombre, apellidos, pass, UO, emailcontact) {

    if (emailcontact == undefined) {
        if (new_email == "" || nombre == "" || apellidos == "" || pass == "" || UO == "") { //Recibe los parametros de la funcion without email
            return false;
        }
    }
    if (new_email == "" || nombre == "" || apellidos == "" || pass == "" || UO == "" || emailcontact == "") { //Recibe los parametros de la funcion with email
        return false;
    }
    return true;
}


///////////////////////////////////////////////////////////////////////////////////////////////////
///Devuelve el numero de columna segun su nombre.                                        
///////////////////////////////////////////////////////////////////////////////////////////////////
function getColIndexByName(colName, sheet) {
  Logger.log("colname getIndex: "+colName);
    var numColumns = sheet.getLastColumn(),
        row = sheet.getRange(1, 1, 1, numColumns).getValues();
    for (var i in row[0]) {
        var name = row[0][i];
        if (name == colName) {
            return parseInt(i) + 1;
        }
    }
    return -1;
}




///////////////////////////////////////////////////////////////////////////////////////////////////
///Genera contraseñas aleatoriamente.                                  
///////////////////////////////////////////////////////////////////////////////////////////////////
function passGenerator() {
    var pass = Math.random().toString(36).substr(2, 9);
    return pass;
}


//////////////////////////////////////////////////////////////////////////////////////////////////////
// Comprueba qué tipo de cuenta tiene el usuario. Si el id de cliente no esta definido quiere decir
// que posee una cuenta personal, no una cuenta de google apps. En ese caso devolveria nulo
// para indicar al usuario que no podrá usar el complemento
//////////////////////////////////////////////////////////////////////////////////////////////////////
function checkAccount() {
    try {
        var lang = getLang(); // Idioma del usuario

        // Valores de las cadenas de texto
        var valores = [
      "No tienes autorización para acceder a este recurso.",
      "Para usar esta aplicación necesitas tener una cuenta de Administrador en un dominio de Google Apps.",
      "Ha ocurrido un error inesperado. Por favor, envíe feedback por medio del apartado Ayuda, para que podamos ayudarle."
    ];

        var mailUser = Session.getActiveUser().getEmail();
        var customerId = AdminDirectory.Users.get(mailUser).customerId;

        return "exito";

    } catch (e) {
        if (e.message == "Not Authorized to access this resource/api") {
            return translate("No tienes autorización para acceder a este recurso.");
        }
        if (e.message == "Resource Not Found: userKey") {
            return translate("Para usar esta aplicación necesitas tener una cuenta de Administrador en un dominio de Google Apps.");
        }

        saveError(e);
        return translate("Ha ocurrido un error inesperado.");
    }
}

/////////////////////////////////////////////////////////////////////////
//Elimina los acentos, las 'ñ' y pasa a minúsculas una cadena de texto.
////////////////////////////////////////////////////////////////////////
function quitaAcentos(str) {
    try {
        str = str.toLowerCase(); //pasa la cadena a minúsculas.
        for (var i = 0; i < str.length; i++) {
            switch (str.charAt(i)) {
                case "ñ":
                    str = str.replace(/ñ/, "n");
                    continue;
                    break;

                case "á":
                    str = str.replace(/á/, "a");
                    continue;
                    break;

                case "é":
                    str = str.replace(/é/, "e");
                    continue;
                    break;

                case "í":
                    str = str.replace(/í/, "i");
                    continue;
                    break;

                case "ó":
                    str = str.replace(/ó/, "o");
                    continue;
                    break;

                case "ú":
                    str = str.replace(/ú/, "u");
                    continue;
                    break;

                default:
                    continue;
            }

        }
        return str;
    } catch (e) {
        saveError(e);
    }
}



///////////////////////////////////////////////////////////////////////////////////////////
// Obtiene el rango de valores de un registro.
///////////////////////////////////////////////////////////////////////////////////////////
function obtener_rango(nombre_columna, sheet) {
    try {
        var index_column = getColIndexByName(nombre_columna, sheet);
        var range = sheet.getRange(2, index_column, sheet.getLastRow() - 1, 1);
        return range;
    } catch (e) {
      var err1 = translate("Las coordenadas o dimensiones del intervalo no son válidas.");
      var err2 = translate("Las coordenadas o dimensiones del intervalo son inválidas.");
     
         if (e.message == err1 || e.message == err2) {
            aviso("No hay datos en la hoja");
            return;
        } else {
            saveError(e);
          aviso("Error inesperado al obtener rango");
            return;
        }
     
    }
}




///////////////////////////////////////////////////////////////////////////////////////////
// Cambia el nombre de la hoja activa
///////////////////////////////////////////////////////////////////////////////////////////
function setNameSheet(activeSheet, namesheet) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(translate(namesheet)),
            ui = SpreadsheetApp.getUi();
        if (sheet) {
            namesheet = ui.prompt('Ya existe una hoja con ese nombre', 'Introduce uno que no exista', ui.ButtonSet.OK).getResponseText();
            setNameSheet(activeSheet, namesheet)
            return;
        } else {
            activeSheet.setName(translate(namesheet));
            return;
        }
    } catch (e) {

    }
}






///////////////////////////////////////////////////////////////////////////////////////////
// Selecciona una hoja y la devuelve.
///////////////////////////////////////////////////////////////////////////////////////////
function getSheetsNames() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets(),
        sheetsNames = [];
    for (var i = 0; i < sheets.length; i++) {
        sheetsNames.push(sheets[i].getName());
    }
    return sheetsNames;
}




///////////////////////////////////////////////////////////////////////////////////////////
// Crea una nueva hoja.
///////////////////////////////////////////////////////////////////////////////////////////
function addSheet(nameSheet, colorRGB) {
  try{
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId(),
      requests = [{
    "addSheet": {
      "properties": {
        "title": nameSheet,
        "gridProperties": {
          "rowCount": 10000,
          "columnCount": 12
        },
        "tabColor": {
          "red": colorRGB[0],
          "green": colorRGB[1],
          "blue": colorRGB[2]
        }
      }
    }
  }];

  var response = Sheets.Spreadsheets.batchUpdate({'requests': requests}, ssId),
      ss = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.flush();
  return ss.getSheetByName(nameSheet);
  }catch(e){
    saveError(e);
    return;
  }
}



///////////////////////////////////////////////////////////////////////////////////////////
// Modifica las propiedades del usuario.
///////////////////////////////////////////////////////////////////////////////////////////
function set_userProperties(name_workSheet, created_users, force_pass){

 var userProperties = PropertiesService.getUserProperties(),
     newProperties = {
       workSheet: name_workSheet,
       createdUsers: created_users,
       forcePass: force_pass
      };
  
 userProperties.setProperties(newProperties);
}





///////////////////////////////////////////////////////////////////////////////////////////
// Obtiene las propiedades del usuario.
///////////////////////////////////////////////////////////////////////////////////////////
function get_userProperties(){
 var userProperties = PropertiesService.getUserProperties(),
     properties=[];
  properties.push(userProperties.getProperty('workSheet'));
  properties.push(userProperties.getProperty('createdUsers'));
  properties.push(userProperties.getProperty('forcePass'));
 return properties;
}


//////////////////////////////////////////////////////////////////////////////////////////////////////////
//devuelve de una cadena la plalabra mas larga.(usado para devolver el nombre mas largo de los apellidos)
// ejemplo: "de la rosa">> return: "rosa" ** n= 0: apell1, n= 1: apell2 **
//////////////////////////////////////////////////////////////////////////////////////////////////////////
function getSplitNames(str, n){
  
  str = str.split(" ");
  str.sort(function(a, b){
   return b.length - a.length;
});
  Logger.log(str[n]);
  return str[n];
}






