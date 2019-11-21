//chrome.storage.sync.clear();location.reload();
function createJSONFile(dataSet) {
    //Agrupar por Cliente
    var groupedByClient = groupByVanillaJS(dataSet, 0);
    //Iterar para crear JSON
    var JSONString = '';
    var JSONCredentialsProd = getOrgsJSONCredentials(groupedByClient);
    JSONString += getJSONString(JSONCredentialsProd);
    createGoogleDriveTextFile(JSONString);
}

function processAllCredentialJSONFile() {
    //Obtener datos de Hojas Origen  Produccion
    var prodSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Produccion");
    var sandSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sandbox");
    //Obetner Longitud Array
    var firstColumnProd = prodSheet.getRange("A2:A").getValues().filter(function(item) {
        return item != ""
    });
    var firstColumnSand = sandSheet.getRange("A2:A").getValues().filter(function(item) {
        return item != ""
    });
    var longArrayProd = firstColumnProd.length + 1;
    var longArraySand = firstColumnSand.length + 1;
    //Crear JSON
    var prodSheetValues = prodSheet.getRange("A2:E" + longArrayProd).getValues();
    var sandSheetValues = sandSheet.getRange("A2:E" + longArraySand).getValues();
    //Marcar P o S
    prodSheetValues.forEach(function(e) {
        e[5] = "P"
    });
    sandSheetValues.forEach(function(e) {
        e[5] = "S"
    });
    var concatSheetValues = prodSheetValues.concat(sandSheetValues);
    createJSONFile(concatSheetValues);
}

function processSelectedCredentialJSONFile() {
    //Determinar si es Sandbox o Produccion
    var sheetName = SpreadsheetApp.getActiveSheet().getName();
    var selectedData = SpreadsheetApp.getSelection().getActiveRange().getValues();
    var selectedRange = SpreadsheetApp.getSelection().getActiveRange().getA1Notation();
    var rangeInitialColumn = selectedRange.slice(0, 1);
    var rangeInitialCell = selectedRange.slice(0, 2);
    var rangeLastColumn = selectedRange.split(":")[1].slice(0, 1);
    var ui = SpreadsheetApp.getUi();
    if (rangeInitialCell == "A1") {
        ui.alert(
            "No puede incluir la celda A1"
        );
    } else if (rangeInitialColumn != "A" || rangeLastColumn != "E") {
        ui.alert(
            "Solo puede seleccionar rangos iniciando en la columna A y terminando en la E"
        );
    } else {
        //Marcar P o S
        if (sheetName.slice(0, 1) == "P") {
            selectedData.forEach(function(e) {
                e[5] = "P"
            });
        } else if (sheetName.slice(0, 1) == "S") {
            selectedData.forEach(function(e) {
                e[5] = "S"
            });
        }
        createJSONFile(selectedData);
    }
}

function getJSONString(JSONCredentials) {
    var JSONString = '{';
    JSONString += '"groups":[';
    JSONString += '{';
    JSONString += '"Name": "General",';
    JSONString += '"isOpen": true,';
    JSONString += '"credentials": []';
    JSONString += '},'
    JSONString += JSONCredentials;
    JSONString += ']';
    JSONString += '}';

    return JSONString;
}


function onOpen() {
    // This line calls the SpreadsheetApp and gets its UI
    // Or DocumentApp or FormApp.
    var ui = SpreadsheetApp.getUi();
    //These lines create the menu items and
    // tie them to functions we will write in Apps Script
    ui.createMenu("Logins")
        .addItem("Crear JSON Total", "processAllCredentialJSONFile")
        .addSeparator()
        .addItem("Crear JSON Selección", "processSelectedCredentialJSONFile")
        .addToUi();
}

function getOrgsJSONCredentials(obj) {

    var JSONString = '';
    for (var a in obj) {
        JSONString += '{';
        JSONString += '"Name":"' + a + '",';
        JSONString += '"isOpen": true,';
        JSONString += '"credentials": [';
        for (i = 0; i < obj[a].length; i++) {
            var endpoint = obj[a][i][5] == 'P' ? 'https://login.salesforce.com/' : 'https://test.salesforce.com/';
            var Id = obj[a][i][5] == 'P' ? 'Production' : 'Sandbox';
            var name = obj[a][i][5] == 'P' ? 'Produccion' : 'Sandbox';
            var descripcion = obj[a][i][5] == 'S' ? obj[a][i][4] : '';
            /*Logger.log(obj[a][i][0]);
            Logger.log(obj[a][i][1]);
            Logger.log(obj[a][i][2]);
            Logger.log(obj[a][i][3]);*/
            JSONString += '{';
            JSONString += '"Name": "' + obj[a][i][0] + ' ' + name + '",';
            JSONString += '"SfName": "' + obj[a][i][1] + '",';
            JSONString += '"Password": "' + obj[a][i][2] + '",';
            JSONString += '"GroupId": "' + obj[a][i][0] + ' ' + name + '",';
            JSONString += '"Description": "' + obj[a][i][0] + ' ' + descripcion + '",';
            JSONString += '"orgId": "",';
            JSONString += '"Type": {';
            JSONString += '"Id": "' + Id + '",';
            JSONString += '"Domain": "' + endpoint + '",';
            JSONString += '"LP": "SETUP"';
            JSONString += '}';
            JSONString += '}';
            JSONString += ',';
        }
        JSONString = JSONString.slice(0, JSONString.length - 1);
        JSONString += ']'; //End Credentials
        JSONString += '},'; //End Group
    }
    JSONString = JSONString.substring(0, JSONString.length - 1);
    return JSONString;
}

/*!
 * Group items from an array together by some criteria or value.
 * (c) 2019 Tom Bremmer (https://tbremer.com/) and Chris Ferdinandi (https://gomakethings.com), MIT License,
 * https://gomakethings.com/a-vanilla-js-equivalent-of-lodashs-groupby-method/
 * @param  {Array}           arr      The array to group items from
 * @param  {String|Function} criteria The criteria to group by
 * @return {Object}                   The grouped object
 */
function groupByVanillaJS(arr, criteria) {
    return arr.reduce(function(obj, item) {

        // Check if the criteria is a function to run on the item or a property of it
        var key = typeof criteria === 'function' ? criteria(item) : item[criteria];

        // If the key doesn't exist yet, create it
        if (!obj.hasOwnProperty(key)) {
            obj[key] = [];
        }

        // Push the value to the object
        obj[key].push(item);

        // Return the object to the next item in the loop
        return obj;

    }, {});
};


/*
 * Origen 
 * https://riptutorial.com/google-apps-script/example/22010/create-a-new-text-file-in-google-root-drive-folder
 */
function createGoogleDriveTextFile(JSONString) {
    var content, fileName, newFile, folder; //Declare variable names
    fileName = "JSON Login Credentials - " + new Date().toString().slice(0, 24); //Create a new file name with date on end
    content = JSONString;
    var blob = Utilities.newBlob('').setDataFromString(content).setContentType("application/json").setName(fileName + ".json");
    //Buscar carpeta
    var folders = DriveApp.getFoldersByName("Logins SFDC JSON");
    Logger.log(folders);
    if (folders.hasNext()) {
        folder = folders.next();
        var folderId = folder.getId();
        newFile = DriveApp.getFolderById(folderId).createFile(blob);
    } else {
        folder = DriveApp.createFolder("Logins SFDC JSON");
        newFile = DriveApp.folder.createFile(blob);
    }
    sendEmail(fileName, newFile, folder);
};

function sendEmail(fileName, newFile, folder) {
    MailApp.sendEmail({
        to: Session.getActiveUser(),
        subject: fileName,
        htmlBody: "Encuentre el archivo en " + newFile.getUrl() + "<br>" +
            "Si desea sobreescribir las credenciales actuales debe eliminar la memoria de la extension usando <br> chrome.storage.sync.clear();location.reload(); <br>" +
            "Ver Manually reset in settings (developer experience) <br> https://www.turnoffthelights.com/support/browser-extension/how-to-reset-the-chrome-extension-settings/",
    });

    var ui = SpreadsheetApp.getUi();
    ui.alert(
        "Correo electronico enviado. Verifique el correo con la ruta para acceder el archivo en su Drive"
    );

}
