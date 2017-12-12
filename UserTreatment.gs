/////////////////////////////////////////////////////////////////////////////////
//Capturar usuarios: devuelve el objeto con todos los datos de cada uno.
////////////////////////////////////////////////////////////////////////////////
function getUsers() {

    var users = AdminDirectory.Users.list({
        customer: "my_customer"
    });
    var obj_users = users.users;
 Logger.log(obj_users[0].name.familyName);
    return obj_users;

}

/////////////////////////////////////////////////////////////////////////////////
// Vuelca los datos de todos los usuarios en la hoja.
////////////////////////////////////////////////////////////////////////////////
function listUsers(sheet_name){
  var workSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  //TRADUCCIONES
  var sheets_names = ['Usuarios_LazySecretary', 'Usuarios creados', 'Errores detectados'];
      sheets_names = translateArray(sheets_names);
  Logger.log(sheets_names);
  Logger.log(sheet_name);
  if(sheet_name != sheets_names[0] && sheet_name != sheets_names[1] && sheet_name != sheets_names[2]){
     generate_header(workSheet)
  }
  var header_values = ['Nombre', 'Apellidos', 'Contraseña', 'Dominio', 'Nuevo email', 'U.O destino', 'Administrador', 'Fecha de creación'];
      header_values = translateArray(header_values);
      header_values.push('V2P', 'Id Googe');
      
      ////////////////////////////////////////////////
  
  
  var obj_users = getUsers(),
      name = [],
      surnames = [],
      pass = [],
      domain = [],
      email = [],
      uo = [],
      isadmin = [],
      creation = [],
      v2p = [],
      id = [];
  
  
  
  for (var i = 0; i < obj_users.length; i++) {
    
      name.push([obj_users[i].name.givenName]);
      surnames.push([obj_users[i].name.familyName]);
      pass.push(["*********"]);
      domain.push([obj_users[i].primaryEmail.split('@')[1]]);
      email.push([obj_users[i].primaryEmail]);
      uo.push([obj_users[i].orgUnitPath]);
      isadmin.push([obj_users[i].isAdmin]);
      creation.push([Utilities.formatDate(new Date(obj_users[i].creationTime), "GMT", "dd-MM-yyyy")]);
      v2p.push([obj_users[i].isEnrolledIn2Sv]);
      id.push([obj_users[i].id]);
  }
  
   workSheet.getRange(2, getColIndexByName(header_values[0], workSheet), name.length, 1).setValues(name);
   workSheet.getRange(2, getColIndexByName(header_values[1], workSheet), name.length, 1).setValues(surnames);
   workSheet.getRange(2, getColIndexByName(header_values[2], workSheet), name.length, 1).setValues(pass);
   workSheet.getRange(2, getColIndexByName(header_values[3], workSheet), name.length, 1).setValues(domain);
   workSheet.getRange(2, getColIndexByName(header_values[4], workSheet), name.length, 1).setValues(email);
   workSheet.getRange(2, getColIndexByName(header_values[5], workSheet), name.length, 1).setValues(uo);
   workSheet.getRange(2, getColIndexByName(header_values[6], workSheet), name.length, 1).setValues(isadmin);
   workSheet.getRange(2, getColIndexByName(header_values[7], workSheet), name.length, 1).setValues(creation);
   workSheet.getRange(2, getColIndexByName(header_values[8], workSheet), name.length, 1).setValues(v2p);
   workSheet.getRange(2, getColIndexByName(header_values[9], workSheet), name.length, 1).setValues(id);
  
  return;
}







///////////////////////////////////////////////////////////////////////////////////////////
// Inserta un usuario en el dominio. Controla varias excepciones
///////////////////////////////////////////////////////////////////////////////////////////
function addUser(email, nombre, apellidos, pass, orgUnit, forcePass) {
    // Valores de las cadenas de texto
    var text_to_translate = [
    "El usuario ya existe",
    '"Admin Directory API" desactivado, se informó al equipo técnico.',
    "No se encuentra la unidad organizativa",
    "No es posible encontrar el dominio especificado. Tienes que introducir el dominio principal de tu cuenta de Google Apps, o un alias de dominio con estado 'activo'.",
    "El correo electrónico del nuevo usuario no es válido",
    "La contraseña tiene menos de 8 caracteres, es demasiado corta.",
    "Error: Es posible que se haya creado el usuario de forma erronea."
  ];

    text_to_translate = translateArray(text_to_translate);


    var user = {
        primaryEmail: email,
        name: {
            givenName: nombre,
            familyName: apellidos
        },
        password: pass,
        changePasswordAtNextLogin: forcePass,
        orgUnitPath: orgUnit
    };
    try {
        user = AdminDirectory.Users.insert(user);
        return "exito";

    } catch (e) {// Tratamiento de errores
      console.log("email incorrecto: "+e.message);
        if (e.message == "Entity already exists.") {
            return text_to_translate[0];
        }
        if (e.message == '"AdminDirectory" is not defined.') {
            saveError(e);
            return text_to_translate[1];
        }
        if (e.message == "Invalid Input: INVALID_OU_ID") {
            return text_to_translate[2];
        }
        if (e.message == "Domain not found.") {
            return text_to_translate[3];
        }
        if (e.message == "Invalid Input: primary_user_email") {
            return text_to_translate[4];
        }
        if (e.message == "Invalid Password") {
            return text_to_translate[5];
        }
      if (e.message == "Backend Error") {
            return text_to_translate[6];
        }
        saveError(e);
        return e.message;
    }
}




/////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Ejecuta la funcion de crear usuarios con o sin email dependiendo de la opción que seleccione el usuario.
// Recibe como parametros el nombre de la hoja con la que se va a trabajar.(string)
////////////////////////////////////////////////////////////////////////////////////////////////////////////
function selectorCreateUsers(name_sheet, click) {
  try{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
      
  //Si la hoja no esxiste avisa.
  if (!ss.getSheetByName(name_sheet)) {
    aviso('No se encuentra la hoja "'+name_sheet+'".');
        return ;
    }
  if(click){//si es la primera llamada, pregunta el email de bienvenida, forzar pass y guarda las propiedades.
    var forcePass = askYesNo("¿Forzar el cambio de contraseña en el primer inicio de sesión?");
        if (forcePass == "cancel") {
          return "cancel";
        }
     
        set_userProperties(name_sheet, 0,  forcePass);
       Utilities.sleep(1000);
        var rsp = createUsersMasive();
        return rsp;
    
  }else{//Si no es la primera llamada, no pregunta forzar pass.
        var rsp = createUsersMasive();
        return rsp;
    
  }
  }catch(e){
    var  usuarios_creados = get_userProperties()[1];
    aviso('Error del servidor.\nSe ha enviado feedback a desarrollo.');
    aviso('Se han creado '+parseFloat(usuarios_creados).toFixed(0)+' usuarios.\nPasa el detector de errores en la hoja de usuarios creados para localizar usuarios no creados.');
    return;
    saveError(e);
  }

}





///////////////////////////////////////////////////////////////////////////////////////////
// Crea usuarios en el dominio a partir de los datos de la hoja (SIN ENVIAR EMAIL)
///////////////////////////////////////////////////////////////////////////////////////////
function createUsersMasive() {
    try {
      //**TRADUCCIONES**//
    var arr_string = ["Uno o más campos necesarios están vacíos", "Usuario creado correctamente", "Usuarios creados", "Crear usuarios", "Estado", "Errores detectados"];
    arr_string = translateArray(arr_string);
    ///////////////////////////
      
        
        var nameWorksheet = get_userProperties()[0],
            workSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameWorksheet),
            usuariosCreados = get_userProperties()[1],
            forcePass = get_userProperties()[2],
            v_estado = [],
            lastLot = false,
            num_registros = workSheet.getLastRow(),
            executedRange = workSheet.getRange(2, 1, 100, workSheet.getLastColumn()),
            createdUsersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(arr_string[2]),
            errorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(arr_string[5]);
      
      //Si quedan menos de 100 registros, es el último lote.
      if(num_registros <=101){
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
                apellidos = registro[1],
                pass = registro[2],
                domain = registro[3],
                new_email = registro[4],
                email_contact = registro[5],
                UO = registro[6];

            // Comprobamos si la fila actual tiene alguno de los campos imprescindibles vacío
            if (!checkData(new_email, nombre, apellidos, pass, UO)) {
                v_estado.push([arr_string[0]]);
                continue;
            };

            // Intentamos crear el usuario
            var estado = addUser(new_email, nombre, apellidos, pass, UO, forcePass);
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
   // createTriggersWithout();// Crea un nuevo disparador
    set_userProperties(nameWorksheet, usuariosCreados, forcePass);
    return "next";
  }else{
  set_userProperties(nameWorksheet, usuariosCreados, forcePass);
    return "success";
  }
        
        return; 
    } catch (e) {// Tratamiento de errores
      var str = ["Las coordenadas o dimensiones del intervalo no son válidas.", "Las coordenadas o dimensiones del intervalo son inválidas."];
          str = translateArray(str);
        if (e.message == str[0] || e.message == str[1]) {
          aviso("No se encuentran datos en la hoja. Asegurate de haber seleccionado la hoja correcta en el menú usuarios.");
            return "failed";
        }
        aviso("¡Error inesperado!\nSe ha enviado feedback al equipo técnico.");
        saveError(e);
        return "failed";
    }
}










////////////////////////////////////////////////////////////////////////////////////////////////
// Detecta los errores en la creacíon de usuarios y los pasa a la hoja de errores detectados.
////////////////////////////////////////////////////////////////////////////////////////////////
function detectErrors(){
   //**TRADUCCIONES**//
    var arr_string = ["Usuarios creados", "Errores detectados", "Estado", "Usuario creado correctamente", 'La hoja de "Usuarios creados" no existe.', 'La hoja de "Errores detectados" no existe.'];
    arr_string = translateArray(arr_string);
    ///////////////////////////
  try{
  var createdUsersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(arr_string[0]),
      errorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(arr_string[1]),
      nameColEstado = arr_string[2],
      str_created = arr_string[3],
      cont_reg = 2,
      cont_errors = 0;
  if(!createdUsersSheet){
  aviso(arr_string[4]);
    return;
  }
  
  if(!errorSheet){
    aviso(arr_string[5]);
    return;
  }
  
 var values_estado = obtener_rango(nameColEstado, createdUsersSheet).getValues();

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
  if(cont_errors == 0){aviso("¡OCURRIÓ ALGO INUSUAL!, ¡No se encontraron errores!"); return;}
  aviso("Se han detectado "+cont_errors+" errores.");
  return;
    
  }catch(e){
    saveError(e);
    return;
  }
}




///////////////////////////////////////////////////////////////////
//Prepara las cuentas de usuarios a partir de nombre y apellidos.
//////////////////////////////////////////////////////////////////
function prepare_users_account(sheet_name) {
    try {
        //sheet_name ="users_LazySecretary";
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
        //*TRANSLATES*//
        var array_translate = ["Dominio no asignado", "Dominio", "Nombre", "Apellidos", "Nuevo email", "Contraseña"];
        array_translate = translateArray(array_translate);
      
        var v_names = obtener_rango(array_translate[2], sheet).getValues(), //captura los nombres en un array
            v_surnames = obtener_rango(array_translate[3], sheet).getValues(), //captura los apellidos en un array
            v_new_email = [],
            v_surnames_split = surnameSplit(v_surnames),
            v_surname1 = [],
            v_surname2 = [],
            v_pass = [],
            username = "",
            new_email = "",
            focusRow = 1,
            duplicateRows = [],
            domain = obtener_rango(array_translate[1], sheet).getValues();
        
      for (var k = 0; k < v_surnames_split.length; k++) {
        
            v_surname1.push(v_surnames_split[k][0]);
       
            v_surname2.push(v_surnames_split[k][1]);
      }
      
      Logger.log("surname1: "+v_surname1[0]);
      Logger.log("surname2: "+v_surname2[0]);
      Logger.log("vnames: "+v_names[0]);
      //Crea el email por defecto y la contraseña.
        for (var i = 0; i < v_names.length; i++) {
            focusRow++;
            var this_domain = domain[i][0].toString()
            //Si el dominio está vacío
            if (this_domain != "") {
                var focus_range = sheet.getRange(i + 2, 1, 1, sheet.getLastColumn()); //Rango de la fila actual.
                username = quitaAcentos(v_names[i].toString().split(" ")[0] + "." + v_surname1[i].toString());
                new_email = username + "@" + domain[i]; //genera el email de la cuenta nueva

                //Si el nombre de usuario que intenta crear ya existe, agrega también el segundo apellido.
                for (var j = 0; j < v_new_email.length; j++) {
                    if (v_new_email[j].toString() == new_email) {
                        duplicateRows.push((j+2));
                        username = quitaAcentos(v_names[i].toString().split(" ")[0].toString() + "." + v_surname1[i].toString()+ "." + v_surname2[i].toString());
                        new_email = username + "@" + domain[i]; //genera el email de la cuenta nueva.
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
        var col_emails = obtener_rango(array_translate[4], sheet),
            col_pass = obtener_rango(array_translate[5], sheet);
        col_emails.setValues(v_new_email); //Agrega los emails a la columna correspondiente.  
        col_pass.setValues(v_pass); //Agrega las contraseñas generadas a la columna correspondiente.  
        sheet.autoResizeColumn(getColIndexByName(array_translate[4], sheet));
      Logger.log("duplic: "+duplicateRows);
      for (var n = 0; n < duplicateRows.length; n++) {
        setNewEmail(duplicateRows[n], sheet);
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




////////////////////////////////////////
//comprueba emails duplicados.
///////////////////////////////////////
function checkDuplicateEmail(){
   var sheet = SpreadsheetApp.getActiveSheet(),
       numRow = 1,
       duplicado = -1;
       col_newEmail = obtener_rango(translate("Nuevo email"), sheet).getValues();
  for (var i = 0; i < col_newEmail.length; i++) {
    numRow++
      for (var j = 0; j < col_newEmail.length; j++) {
        if(col_newEmail[i][0] == col_newEmail[j][0]){
          duplicado++;
          
        }
        if(duplicado > 0){Logger.log(col_newEmail[i][0]+" tiene "+duplicado+" duplicados.");}
      }
   
    duplicado = -1;
  }
  
}



///////////////////////////////////////////////////////////////////
//Modifica el email de la fila indicada.
//////////////////////////////////////////////////////////////////
function setNewEmail(row, sheet){
   //*TRANSLATES*//
        var array_translate = ["Dominio", "Nombre", "Apellidos", "Nuevo email"];
        array_translate = translateArray(array_translate);
  
  var rangeName = sheet.getRange(row, getColIndexByName(array_translate[1], sheet), 1, 1).getValue(),
      domain = sheet.getRange(row, getColIndexByName(array_translate[0], sheet)).getValue(),
      nombre = quitaAcentos(rangeName.split(" ")[0]),
      rangeApell = [];
      rangeApell.push(sheet.getRange(row, getColIndexByName(array_translate[2], sheet), 1, 1).getValue());
  var v_apells= surnameSplit(rangeApell),
      newEmail = (nombre.toLowerCase()+"."+v_apells[0][0].toLowerCase()+"."+v_apells[0][1].toLowerCase()+"@"+domain);
     sheet.getRange(row, getColIndexByName(array_translate[3], sheet)).setValue(newEmail);
  return;
}





///////////////////////////////////////////////////////////////////
//Separa los apellidos eliminando los conectores.
//////////////////////////////////////////////////////////////////
function surnameSplit(v_surname){
 var v_apell = [],
     v_apells = [];
  for (var j = 0; j < v_surname.length; j++) {
   var v_valid_apell = [];
     v_apell = v_surname[j].toString().split(" ");
   for (var i = 0; i < v_apell.length; i++) {
     //Comprueba si es un conector.
     if(v_apell[i].length < 4 || v_apell[i].toLowerCase() == "santa" || v_apell[i].toLowerCase() == "Della" || v_apell[i].toLowerCase() == "Van't" || v_apell[i].toLowerCase() == "Aus'm"){
       continue;
     }
     v_valid_apell.push(v_apell[i]);
     continue;
   }
    v_apells.push(v_valid_apell);
  }
 return v_apells;
}





function sendCredentials(){

  
  
}














