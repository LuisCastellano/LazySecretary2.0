// Cuando se instala el documento se añaden las opciones del complemento al menu
function onInstall(e) {

    var ui = SpreadsheetApp.getUi();
    ui.createAddonMenu()
        .addItem(translate('Iniciar'), 'getStarted')
        .addToUi();
}

// Cuando se abre el documento se añaden las opciones del complemento al menu
function onOpen(e) {
    var ui = SpreadsheetApp.getUi();
    ui.createAddonMenu()
        .addItem(translate('Iniciar'), 'getStarted')
        .addToUi();
}


// Abre una sideBar para guiar al usuario en el proceso de creación de usuarios 
function getStarted() {
    // Comprueba que tiene una cuenta de Google Apps como administrador.
    var typeAccount = checkAccount();
    if (typeAccount != "exito") {
        aviso(typeAccount);
        return;
    }
  
  init_load();

    // Creamos el sidebar
    var html = HtmlService.createHtmlOutputFromFile('SideBar')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setTitle('LazySecretary');

    // Lo mostramos
    SpreadsheetApp.getUi()
        .showSidebar(html);
}


function init_load(){
  
// Display a modal dialog box with custom HtmlService content.
 var html = HtmlService.createHtmlOutputFromFile('init_loading')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(300)
        .setHeight(150);
 SpreadsheetApp.getUi().showModalDialog(html,' ');
}