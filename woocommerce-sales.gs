var ss = SpreadsheetApp.openByUrl("La url de tu archivo Google Sheets");
var sheet = ss.getSheetByName('Nombre de la pestaña que vamos a trabajar en el Google Sheets');
var fechaMin = sheet.getRange('b25').getValue(); // Acá fijamos la fecha de inicio para el request
var fechaMax = sheet.getRange('c25').getValue(); // Acá fijamos la fecha de término para el request


function Obtener_sales(){
var ck = "ck_Tu-Llave";
var cs = "cs_Tu-Llave";
var website = "https://Tu-sitio-web.com";

let url = website + "/wp-json/wc/v3/reports/sales?date_min=" + fechaMin + "&date_max=" + fechaMax; 


let encoded = Utilities.base64EncodeWebSafe(ck + ':' + cs, Utilities.Charset.UTF_8);

let options = {
    "muteHttpExceptions":true,
    "headers": {
        "Authorization": "Basic " + encoded,
    }
};

let result = UrlFetchApp.fetch(url, options);
result = JSON.parse(result);
var celda = sheet.getRange("c26"); //El resultado va a quedar en la celda c26
  
  celda.setValue(result);
}
