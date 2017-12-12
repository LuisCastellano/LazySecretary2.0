/////////////////////////////////////////////////////////////////////////////////////////////////////
// Genera una cabecera con los diferentes valores a rellenar en la página de creación de usuarios
////////////////////////////////////////////////////////////////////////////////////////////////////
/*function generate_header_addUser(workSheet) {
    try {
        var header_values = ['Nombre', 'Primer apellido', 'Segundo apellido', 'Contraseña', 'Dominio', 'Nuevo email', 'Email de contacto', 'U.O destino', 'Estado'];
        header_values = translateArray(header_values);
        for (var i = 0; i < header_values.length; i++) {
            workSheet.getRange(1, (i + 1)).clearDataValidations().setValue(header_values[i]).setBackground('#03a9f4').setFontColor('#fff').setFontSize(15).setHorizontalAlignment('center').setWrap(true);
            workSheet.autoResizeColumn(i + 1); //Adapta el tamaño de la columna al texto de la cabecera. 
        }


        workSheet.setFrozenRows(1);
        build_rule_dataV_uo(workSheet);
        build_rule_dataV_domain(workSheet);
    } catch (e) {
        saveError(e);
    }

}
*/









///////////////////////////////////////////////////////////////////////////////////////////
// Crea usuarios en el dominio a partir de los datos de la hoja (ENVIANDO EMAIL)
///////////////////////////////////////////////////////////////////////////////////////////
/*function createUsersWithEmail(sheet) {
    var actualRow = 1, // para saber la fila por donde vamos.
        colorR = "#e57373",
        colorG = "#81c784",
        usuariosCreados = 0; // para llevar un control de los usuarios creados




    var data = getData(); // Recoge los datos de la hoja de calculo
    if (data == null) {
        return "error";
    }

    var registro = data[i],
        nombre = registro[0],
        apll1 = registro[1],
        apll2 = registro[2],
        pass = registro[3],
        domain = registro[4],
        new_email = registro[5],
        email_contact = registro[6],
        UO = registro[7];



    // Valores de las cadenas de texto
    var valores = [
    "Los campos obligatorios no pueden estar vacíos",
    "Usuario no creado: introduce un correo electrónico válido para enviar el correo de bienvenida",
    "Usuario no creado: has agotado todos los correos electrónicos que puedes enviar en un día",
    "El usuario se ha creado, correo electrónico enviado",
    "El usuario se ha creado. Configurado para no enviar correo electrónico de bienvenida"
  ];
    for (var i = 0; i < valores.length; i++) {
        valores[i] = translate(valores[i]);
    }

    for (var i = 0; i < data.length; i++) {

        var sheet = SpreadsheetApp.getActiveSheet(), // Hoja actual
            decision = true, // Indicara si seguir o no en caso de permitir el envio de correo de bienvenida sin tener suficiente cuota
            cuotaEmails = checkDailyQuota();
        // Si tenemos menos cuota que usuarios vamos a crear
        if (cuotaEmails < data.length) {
            aviso("No tienes suficientes correos electrónicos para enviar mensajes de bienvenida a todos los usuarios.\nCorreos disponibles: " + checkDailyQuota() + "\nSe canceló la creación de usuarios.");
            return;
        }
        // Si tenemos los mismos usuarios a crear que cuota de emails diaria.
        else if (cuotaEmails == data.length) {
            decision = askYesNo("Vas a utilizar todos tus correos electrónicos diarios disponibles. ¿Quieres continuar?");
        }

        // Si se decide no seguir se muestra un mensaje indicando que no se han creado usuarios
        if (decision == false) {
            aviso("Se canceló la creación de usuarios.");
            return; // Termina el programa
        }


        // Intentamos crear el usuario
        var estado = addUser(new_email, nombre, apll1 + " " + apll2, pass, UO);

        if (estado == "exito") {
            sendEmail(email_contact, new_email, pass, nombre, domain);
            set_cell_estado(actualRow, valores[3], colorG, sheet); // usuario creado y email enviado

            usuariosCreados++;
        } else {
            set_cell_estado(actualRow, estado, colorR, sheet);
        }
    }

}*/


/*function addSheet(spreadsheetId) {
  SpreadsheetApp.get
  var requests = [{
    "addSheet": {
      "properties": {
        "title": "Hoja nuevad2",
        "gridProperties": {
          "rowCount": 10000,
          "columnCount": 12
        },
        "tabColor": {
          "red": 1.0,
          "green": 0.3,
          "blue": 0.4
        }
      }
    }
  }];

  var response =
      Sheets.Spreadsheets.batchUpdate({'requests': requests}, spreadsheetId);
  Logger.log("Created sheet with ID: " +
      response.replies[0].addSheet.properties.sheetId);
}*/










/*function ppp(){
  set_userProperties("CAMBIOO");
  
  return;
}




function set_userProperties(name_workSheet){
 var userProperties = PropertiesService.getUserProperties(),
     newProperties = {
       workSheet: name_workSheet
      };
  
 userProperties.setProperties(newProperties);
}


function get_userProperties(){
 var userProperties = PropertiesService.getUserProperties();
    
  
 Logger.log(userProperties.getProperty('workSheet'));
}
*/




function ff(){
  var forcePass = Boolean(1);

  Logger.log(forcePass);
}

/*
<div id="lbl_pag_create_cont">
            <ul>
                <li><label id="lbl_create1"></label></li>
                <li><label id="lbl_create2"></label></li>
                <li><label id="lbl_create3"></label></li>
            </ul>
        </div>
*/



////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////
/**
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
         }/
        return obj_orgUnits;
    } catch (e) {
        if (e.message == 'Cannot call method "sort" of undefined.') {
            aviso(translate("No se encontró ninguna U.O."));
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



/////////////////////////////////////////////////////////////////////////////////
//Capturar usuarios: devuelve el objeto con todos los datos de cada uno.
////////////////////////////////////////////////////////////////////////////////
function getUsers() {

    var users = AdminDirectory.Users.list({
        customer: "my_customer"
    });
    var obj_users = users.users;
    return obj_users;

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



///////////////////////////////////////////////////////////////////////////////////////////
// Inserta un usuario en el dominio. Controla varias excepciones
///////////////////////////////////////////////////////////////////////////////////////////
function addUser(email, nombre, apellidos, pass, orgUnit) {
    // Valores de las cadenas de texto
    var text_to_translate = [
    "El usuario ya existe",
    'Debes activar "Admin Directory API" en este proyecto',
    "No se encuentra la unidad organizativa",
    "No es posible encontrar el dominio especificado. Tienes que introducir \n" + "el dominio principal de tu cuenta de Google Apps, o un alias de dominio  \n" + "con estado 'activo'.",
    "El correo electrónico del nuevo usuario no es válido",
    "La contraseña tiene menos de 8 caracteres, es demasiado corta.",
    "Ha ocurrido un error inesperado. Por favor, envíe feedback por medio del apartado Ayuda, para que podamos ayudarle."
  ];

    text_to_translate = translateArray(text_to_translate);


    var user = {
        primaryEmail: email,
        name: {
            givenName: nombre,
            familyName: apellidos
        },
        password: pass,
        orgUnitPath: orgUnit
    };
    try {
        user = AdminDirectory.Users.insert(user);
        return "exito";

    } catch (e) {// Tratamiento de errores
        if (e.message == "Entity already exists.") {
            return text_to_translate[0];
        }
        if (e.message == '"AdminDirectory" is not defined.') {
            return text_to_translate[1];
        }
        if (e.message == "Invalid Input: INVALID_OU_ID") {
            return text_to_translate[2];
        }
        if (e.message == "Resource Not Found: domain") {
            return text_to_translate[3];
        }
        if (e.message == "Invalid Input: primary_user_email") {
            return text_to_translate[4];
        }
        if (e.message == "Invalid Password") {
            return text_to_translate[5];
        }
        saveError(e);
        return text_to_translate[6];
    }
}





//////////////////////////////////////////////////////////////////////////////////////////
//Crea un disparador para ejecutar la funcion de crear usuarios cada 5 minutos.
/////////////////////////////////////////////////////////////////////////////////////////
function createTriggersWithout() {
    ScriptApp.newTrigger('createUsersWithoutEmail').timeBased().everyMinutes(1).create();
}




//////////////////////////////////////////////////////
//Elimina el trigger WithoutEmail.
/////////////////////////////////////////////////////
function deleteTriggersWithout() {
    var triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActiveSpreadsheet());
    for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() == 'createUsersWithoutEmail') {
            ScriptApp.deleteTrigger(triggers[i]);
            return;
        }
        continue;
    }
}







/////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Ejecuta la funcion de crear usuarios con o sin email dependiendo de la opción que seleccione el usuario.
// Recibe como parametros el nombre de la hoja con la que se va a trabajar.(string)
////////////////////////////////////////////////////////////////////////////////////////////////////////////
function selectorCreateUsers(name_sheet) {
  set_userProperties(name_sheet, 0);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(name_sheet)) {
    aviso('No se encuentra la hoja "'+name_sheet+'".');
        return ;
    }
    var workSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name_sheet),
        welcomeMail = askYesNo("¿Desea enviar un email de bienvenida a los nuevos usuarios?");
        
    if (welcomeMail) {
        var rsp = createUsersWithEmail();
        return rsp;
    } else {
        var rsp = createUsersWithoutEmail();
        return rsp;
    }

}









///////////////////////////////////////////////////////////////////////////////////////////
// Crea usuarios en el dominio a partir de los datos de la hoja (SIN ENVIAR EMAIL)
///////////////////////////////////////////////////////////////////////////////////////////
function createUsersWithoutEmail() {
    try {
    deleteTriggersWithout();// Elimina el disparador creado anteriormente.
      //**TRADUCCIONES//
    var arr_string = ["Uno o más campos necesarios están vacíos", "Usuario creado correctamente", "Usuarios creados", "Crear usuarios", "Estado", "Errores detectados"];
    arr_string = translateArray(arr_string);
    ///////////////////////////
      
        
        var nameWorksheet = get_userProperties()[0],
            workSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameWorksheet),
            usuariosCreados = get_userProperties()[1],
            v_estado = [],
            lastLot = false,
            num_registros = workSheet.getLastRow(),
            executedRange = workSheet.getRange(2, 1, 100, workSheet.getLastColumn()),
            createdUsersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(arr_string[2]),
            errorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(arr_string[5]);
            
      //Si quedan menos de 100 registros, es el último lote.
      if(num_registros <=101){
        Logger.log("Entra en lote final");
        workSheet.insertRowAfter(workSheet.getLastRow());// Añade una fila al final para asegurar que no se queda sin filas a la hora de eliminar el último lote-
        executedRange = workSheet.getRange(2, 1, (num_registros-1), workSheet.getLastColumn());//Captura el ultimo lote
        lastLot = true;// Avtiva el ultimo lote para NO crear el trigger y acabar la ejecución.
      }

        
        var data = executedRange.getValues(); // Recoge el rango
        if (data == null) {
          aviso("La hoja está vacía.");
            return "failed";
        }

        var rangecol_estado = workSheet.getRange(2, getColIndexByName(arr_string[4], workSheet), data.length, 1);// Rango de 100 filas en la columna estado.

        for (var i = 0; i < data.length; i++) {
            var registro = data[i],
                nombre = registro[0],
                apll1 = registro[1],
                apll2 = registro[2],
                pass = registro[3],
                domain = registro[4],
                new_email = registro[5],
                email_contact = registro[6],
                UO = registro[7],
                apellidos = apll1 + " " + apll2;

            // Comprobamos si la fila actual tiene alguno de los campos imprescindibles vacío
            if (!checkData(new_email, nombre, apellidos, pass, UO)) {
                v_estado.push([arr_string[0]]);
                continue;
            };

            // Intentamos crear el usuario
            var estado = addUser(new_email, nombre, apellidos, pass, UO);
            if (estado == "exito") {
                usuariosCreados++;
                v_estado.push([arr_string[1]]);
                continue;
            } else {
                v_estado.push([estado]);
                continue;
            }
        }
        
        rangecol_estado.setValues(v_estado);//Vuelca el array de estados en la columna "Estado".
        Utilities.sleep(2000);//Duerme la ejecución 2 seg.
        copyRange(workSheet, executedRange, createdUsersSheet);//Copia el rango a la hoja de usuarios creados.
        Utilities.sleep(3000);//Duerme la ejecución 3 seg.
        workSheet.deleteRows(2, data.length);
        SpreadsheetApp.flush();
        Utilities.sleep(2000);//Duerme la ejecución 2 seg.
        
  
  if(!lastLot){//si NO es el ultimo lote
    createTriggersWithout();// Crea un nuevo disparador
    set_userProperties(nameWorksheet, usuariosCreados);
    return;
  }else{
  set_userProperties(nameWorksheet, usuariosCreados);
    return "sucess";
  }
        
        return; //["finOK", usuariosCreados];
    } catch (e) {// Tratamiento de errores
        if (e.message == "Las coordenadas o dimensiones del intervalo no son válidas.") {
          aviso("No se encuentran datos en la hoja.");
            return "failed";
        }
        aviso("¡Error inesperado!\nSe ha enviado feedback al equipo técnico.\n" + e.message+e.lineNumber+e.fileName);
        saveError(e);
        return "failed";
    }
}




///////////////////////////////////////////////////////////////////////////////////////////
// Detecta los errores y los pasa a la hoja de errores detectados.
///////////////////////////////////////////////////////////////////////////////////////////
function detectErrors(createdUsersSheet, errorSheet, nameColEstado, str_created){
  createdUsersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios creados");
  errorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Errores detectados");
  nameColEstado = "Estado";
  str_created = "Usuario creado correctamente";
  var values_estado = obtener_rango(nameColEstado, createdUsersSheet).getValues(),
      cont_reg = 2,
      cont_errors = 0;

  for (var i = 0; i < values_estado.length; i++) {
    
    if(values_estado[i].toString() != str_created){
      var errorRange = createdUsersSheet.getRange((i+cont_reg), 1, 1, createdUsersSheet.getLastColumn());    
      copyRange(createdUsersSheet, errorRange, errorSheet);
      createdUsersSheet.deleteRow((i+cont_reg));
      cont_reg--;
      cont_errors++;
      continue
    }
    continue;
  }
  if(cont_errors == 0){aviso("¡Ocurrió algo inusual!\n¡No se encontraron errores!\n<<EXITO>>"); return;}
  aviso("Se han detectado "+cont_errors+" errores.");
  return;
}



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Prepara la hoja de trabajo para crear usuarios.(recibe el nombre de la hoja seleccionada en el desplegable)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function prepareWorkSheets(name_sheet) {
  try{
    name_sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  var response = askYesNo("Has seleccionado la hoja '"+name_sheet+"' para la creación de nuevas cuentas.\nAl preparar la hoja se eliminarán todos los datos de esta.\n¿Estás seguro?");
  if(response){
    var workSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name_sheet);
    workSheet.setTabColor('#29b6f6');
    workSheet.clear();
    generate_header_addUser(workSheet);
     //**TRADUCCIONES//
    var arr_string = ["Usuarios creados", "Errores detectados"];
    arr_string = translateArray(arr_string);
    ///////////////////////////
    Utilities.sleep(1000);
   var createdSheet = getOrCreateSheet(arr_string[0], [102,187,106], workSheet);// crea la hoja de usuarios creados si no existe - rgb(102, 187, 106).Green
       Utilities.sleep(1000);
   var errorSheet = getOrCreateSheet(arr_string[1], [239,83,80], workSheet);// crea la hoja de errores detectados si no existe - rgb(239, 83, 80). Red
    Utilities.sleep(1000);
    SpreadsheetApp.flush();
    return "succesful";
  }
  else{return "cancel";}
    
  }catch(e){
    saveError(e);
    return "failed"
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
// Genera una cabecera con los diferentes valores a rellenar
///////////////////////////////////////////////////////////////////////////////////////////
function generate_header_addUser(workSheet) {
    try {
        var header_values = ['Nombre', 'Primer apellido', 'Segundo apellido', 'Contraseña', 'Dominio', 'Nuevo email', 'Email de contacto', 'U.O destino', 'Estado'];
        header_values = translateArray(header_values);
        for (var i = 0; i < header_values.length; i++) {
            workSheet.getRange(1, (i + 1)).clearDataValidations().setValue(header_values[i]).setBackground('#03a9f4').setFontColor('#fff').setFontSize(15).setHorizontalAlignment('center').setWrap(true);
            workSheet.autoResizeColumn(i + 1); //Adapta el tamaño de la columna al texto de la cabecera. 
        }


        workSheet.setFrozenRows(1);
        build_rule_dataV_uo(workSheet);
        build_rule_dataV_domain(workSheet);
    } catch (e) {
        saveError(e);
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
    //rellena el array con las rutas de la uo.
    for (var i = 0; i < obj_orgUnits.length; i++) {
        v_uo.push(obj_orgUnits[i].orgUnitPath);
    }
    var range_DV = sheet.getRange(2, colUO, sheet.getMaxRows()),
        //Crea la regla del dataValidation.
        rule_action = SpreadsheetApp.newDataValidation().requireValueInList(v_uo).build();
    range_DV.setDataValidation(rule_action); //Agrega el DataValidation al rango.

}






///////////////////////////////////////////////////////////////////
//Prepara las cuentas de usuarios a partir de nombre y apellidos.
//////////////////////////////////////////////////////////////////
function prepare_users_account(sheet) {
    try {
        sheet = SpreadsheetApp.getActiveSheet();
        //*TRANSLATES//
        var array_translate = ["Dominio no asignado", "Dominio", "Nombre", "Primer apellido", "Segundo apellido", "Nuevo email", "Contraseña"];
        array_translate = translateArray(array_translate);

        var v_names = obtener_rango(array_translate[2], sheet).getValues(), //captura los nombres en un array
            v_surname1 = obtener_rango(array_translate[3], sheet).getValues(), //captura los apellidos en un array
            v_surname2 = obtener_rango(array_translate[4], sheet).getValues(), //captura los apellidos en un array
            v_new_email = [],
            v_pass = [],
            username = "",
            new_email = "",
            focusRow = 1,
            domain = obtener_rango(array_translate[1], sheet).getValues();
        //Crea el email por defecto y la contraseña.
        for (var i = 0; i < v_names.length; i++) {
            focusRow++;
            var this_domain = domain[i][0].toString()
            //Si el dominio está vacío
            if (this_domain != "") {
                var focus_range = sheet.getRange(i + 2, 1, 1, sheet.getLastColumn()); //Rango de la fila actual.
                username = quitaAcentos(v_names[i].toString().split(" ")[0] + "." + v_surname1[i].toString().split(" ")[0]);
                new_email = username + "@" + domain[i]; //genera el email de la cuenta nueva

                //Si el nombre de usuario que intenta crear ya existe, agrega también el segundo apellido.
                for (var j = 0; j < v_new_email.length; j++) {
                    if (v_new_email[j].toString() == new_email) {
                        username = quitaAcentos(v_names[i].toString().split(" ")[0] + "." + v_surname1[i].toString().split(" ")[0] + "." + v_surname2[i].toString().split(" ")[0]);
                        new_email = username + "@" + domain[i]; //genera el email de la cuenta nueva
                    }
                }
                v_new_email.push([new_email]); //almacena los emails de usuario en un array de objetos
                var pwd = passGenerator();
                v_pass.push([pwd]);
                continue;
            } else {
                //Se rellena los arrays sin valores para que no descuadre la longitud.
                v_pass.push([" "]);
                v_new_email.push([array_translate[0]]);
                //////////////////////////
                var actualRange = sheet.getRange(focusRow, getColIndexByName(array_translate[1], sheet));
                continue;
            }

        }
        var col_emails = obtener_rango(array_translate[5], sheet),
            col_pass = obtener_rango(array_translate[6], sheet);
        col_emails.setValues(v_new_email); //Agrega los emails a la columna correspondiente.  
        col_pass.setValues(v_pass); //Agrega las contraseñas generadas a la columna correspondiente.  
        sheet.autoResizeColumn(getColIndexByName(array_translate[5], sheet));
        SpreadsheetApp.flush();
    } catch (e) {
        if (e.message == "The coordinates or dimensions of the range are invalid.") {
            aviso("No hay datos en la hoja: \n" + e.fileName + "\n" + e.lineNumber);
            return;
        } else {
            saveError(e);
            return;
        }
    }

}*/




//////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////







function surnameSplisst(v_surname){
  v_surname = ["de la Rosa Martinez", "de las Marismas Rodriguez"];
 var v_conn = ["Am", "Aus'm", "Vom", "Zum", "Zur", "La", "Le", "Du", "Des", "D'", "De", "Ver", "Del", "Den",
   "op de", "ter", "Van", "Der", "ten", "Van't", "A", "Da", "Della", "Di", "Li", "Lo", "Von",
   "Do", "Das", "Dos", "Los", "Las", "Les", "Y", "San", "I", "Mac", "Mc", "Santa"],
     v_split = [],
     d =[],
     ap1 = "",
     ap2 = "";
  
  for (var i = 0; i < v_surname.length; i++) {
    v_split = v_surname[i].split(" ");
    for (var j = 0; j < v_split.length; j++) {
      for (var k = 0; k < v_conn.length; k++) {
      if (v_split[j] == v_conn[k] || v_split[j] == v_conn[k].toLowerCase()|| v_split[j] == v_conn[k].toUpperCase()){
       ap1+=v_split[j];
        
      }else{
        ap1+=v_split[j];
        ap2+=v_split[j];
      }
      }
      
    }
  }
  
}




function names(){

var nombre = "de las marismas blanca paloma rocio del mar";//Nombre completo en string que queremos separar.
var arregloNombre = nombre.split(' ');//Array del nombre con las palabras separadas en cada posición.
var fullName = [];//Array que contendrá el nombre final.
var palabrasReservadas =['da', 'de', 'del', 'la', 'las', 'los', 'san', 'santa'];//Palabras de apellidos y nombres compuestos, aquí podemos agregar más palabras en caso de ser necesario.
var auxPalabra = "";//Variable auxiliar para concatenar los apellidos compuestos.
arregloNombre.forEach(function(name){//Iteramos el array del nombre.
var  nameAux = name.toLowerCase();//convertimos en minúscula la palabra que se esta iterando para poder hacer la búsqueda de esta palabra en nuestro arreglo de "palabrasReservadas".
    var ap1="";
      var ap2="";
  if(palabrasReservadas.indexOf(nameAux)!=-1)//Cuando la palabra existe dentro de nuestro array, la funcion "indexOf" nos arrojara un numero diferente de -1.
 {
 auxPalabra += name+' ' ;//Concatenamos y guardamos en nuestra variable auxiliar la palabra detectada.
 }
 else {//En caso de que la palabra no existe en nuestro array de palabras reservadas, hacemos un push a la variable "fullName" que contendrá el nombre final
 fullName.push(auxPalabra+name);
 auxPalabra = "";//Limpiamos la variable auxiliar
 }
 });
//Al final de la iteración vamos a tener un array en el cual la posicion 0 y 1 contienen los apellidos Paterno y Materno respectivamente.
//las siguientes posiciones despues de eso contendra el nombre
ap1=fullName[0];
ap2=fullName[1];  

delete fullName[0];//Eliminamos la posición del apellido paterno
delete fullName[1];//Eliminamos la posición del apellido materno
nombreCompleto = "";//Variable que contiene el puro nombre
fullName.forEach(function(nombre){//Iteramos en caso de que la persona tenga un nombre compuesto, ejemplo: Juan Manuel 
 if(nombre!="")
 {
 nombreCompleto +=nombre+" ";//Concatenamos el nombre
 }
});
Logger.log("Nombre completo: "+nombreCompleto);//Nombre completo sin apellidos
  Logger.log("Apellido paterno: "+fullName[0]);//Apellido Paterno
  Logger.log("Apellido materno: "+fullName[1]);//Apellido Materno


}


///////////////////////////////////////////////////////////////////////////////////////////
// Genera una cabecera con los diferentes valores a rellenar en la página principal
///////////////////////////////////////////////////////////////////////////////////////////
function generate_header2(workSheet) {
    try {
      //TRADUCCIONES
        var header_values = ['Nombre', 'Primer apellido', 'Segundo apellido', 'Contraseña', 'Dominio', 'Nuevo email', 'Email de contacto', 'U.O destino', 'Estado', 'Administrador', 'Fecha de creación'],
            fistSheetName = translate('Usuarios')+'_LazySecretary';
            header_values.push('V2P', 'Id Googe');
      header_values = translateArray(header_values);
      
      ////////////////////////////////////////////////
      
      //Añade los valores de la cabecera a las celdas
        for (var i = 0; i < header_values.length; i++) {
             workSheet.getRange(1,i+1).setValue(header_values[i]);
        }
      //Obtiene el numero correspondiente a la columna.
      var colEstado = getColIndexByName(header_values[8], workSheet),
          colIdGoogle = getColIndexByName(header_values[12], workSheet);
      
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
        workSheet.getRange(1,getColIndexByName(header_values[11], workSheet)).setNote(translate("Verificación en dos pasos"));
        
      //////////////////////
        workSheet.setFrozenRows(1);
      if(workSheet.getName() == fistSheetName){
        build_rule_dataV_uo(workSheet);
        build_rule_dataV_domain(workSheet);
      }
        SpreadsheetApp.flush();
        return;
        
    } catch (e) {
        saveError(e);
      return;
    }

}







///////////////////////////////////////////////////////////////////
//Prepara las cuentas de usuarios a partir de nombre y apellidos.
//////////////////////////////////////////////////////////////////
function prepare_users_account2 (sheet_name) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
        //*TRANSLATES*//
        var array_translate = ["Dominio no asignado", "Dominio", "Nombre", "Primer apellido", "Segundo apellido", "Nuevo email", "Contraseña"];
        array_translate = translateArray(array_translate);
      
        var v_names = obtener_rango(array_translate[2], sheet).getValues(), //captura los nombres en un array
            v_surname1 = obtener_rango(array_translate[3], sheet).getValues(), //captura los apellidos en un array
            v_surname2 = obtener_rango(array_translate[4], sheet).getValues(), //captura los apellidos en un array
            v_new_email = [],
            v_pass = [],
            username = "",
            new_email = "",
            focusRow = 1,
            duplicateRows = [],
            domain = obtener_rango(array_translate[1], sheet).getValues();
        //Crea el email por defecto y la contraseña.
        for (var i = 0; i < v_names.length; i++) {
            focusRow++;
            var this_domain = domain[i][0].toString()
            //Si el dominio está vacío
            if (this_domain != "") {
                var focus_range = sheet.getRange(i + 2, 1, 1, sheet.getLastColumn()); //Rango de la fila actual.
                username = quitaAcentos(v_names[i][0].split(" ")[0] + "." + getSplitNames(v_surname1[i][0]));
                new_email = username + "@" + domain[i]; //genera el email de la cuenta nueva

                //Si el nombre de usuario que intenta crear ya existe, agrega también el segundo apellido.
                for (var j = 0; j < v_new_email.length; j++) {
                    if (v_new_email[j].toString() == new_email) {
                      Logger.log(j+2);
                        duplicateRows.push((j+2));
                        username = quitaAcentos(v_names[i][0].split(" ")[0] + "." + getSplitNames(v_surname1[i][0]) + "." + getSplitNames(v_surname2[i][0]));
                        new_email = username + "@" + domain[i]; //genera el email de la cuenta nueva
                    }
                }
                v_new_email.push([new_email]); //almacena los emails de usuario en un array de objetos
                var pwd = passGenerator();
                v_pass.push([pwd]);
                continue;
            } else {
                //Se rellena los arrays sin valores para que no descuadre la longitud.
                v_pass.push([" "]);
                v_new_email.push([array_translate[0]]);
                //////////////////////////
                var actualRange = sheet.getRange(focusRow, getColIndexByName(array_translate[1], sheet));
                continue;
            }

        }
        var col_emails = obtener_rango(array_translate[5], sheet),
            col_pass = obtener_rango(array_translate[6], sheet);
        col_emails.setValues(v_new_email); //Agrega los emails a la columna correspondiente.  
        col_pass.setValues(v_pass); //Agrega las contraseñas generadas a la columna correspondiente.  
        sheet.autoResizeColumn(getColIndexByName(array_translate[5], sheet));
      for (var k = 0; k < duplicateRows.length; k++) {
        setNewEmail(duplicateRows[k]);
      }
        SpreadsheetApp.flush();
      aviso("Se han preparado las cuentas de usuario; Quitado acentos, mayúsculas, creado emails, reducido al mínimo el nº de duplicados y generado contraseñas válidas.");
      return;
      
    } catch (e) {
      var err1 = translate("Las coordenadas o dimensiones del intervalo no son válidas.");
      var err2 = translate("Las coordenadas o dimensiones del intervalo son inválidas.");
       if (e.message == err1 || e.message == err2) {
            aviso("No hay datos en la hoja");
            return;
        } else {
            saveError(e);
            return;
        }
    }

}






