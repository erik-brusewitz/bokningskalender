//Skrivet av Erik "Smulan" Brusewitz, skriv om du har nån fråga

//Skicka meddelanden i spreadsheetet:
//SpreadsheetApp.getUi().alert("Meddelande: " + );


//Returnerar namnet på givet datums månad
function monthName(date) {
  var numMonth = date.getMonth();
  switch (numMonth) {
    case 0: return "Januari";
    case 1: return "Februari";
    case 2: return "Mars";
    case 3: return "April";
    case 4: return "Maj";
    case 5: return "Juni";
    case 6: return "Juli";
    case 7: return "Augusti";
    case 8: return "September";
    case 9: return "Oktober";
    case 10: return "November";
    case 11: return "December";
  }
}

//Input namnet på en kolumn, returnerar indexet på kolumnen, indexering från 0.
function getColIdx(firstRow, name){
  return firstRow.indexOf(name);
}

//Skickar ett mail till "email_address" att bokningen misslyckades med meddelandet "message"
function sendFailedEmail(email_address, message) {
  GmailApp.sendEmail(email_address, "Bokning av Helikoptern misslyckades", message);
}

//Skickar ett mail till "email_address" att bokningen lyckades med meddelandet "message"
function sendSuccessfulEmail(email_address, message) {
  GmailApp.sendEmail(email_address, "Bokningsbekräftelse Helikoptern", message);
}

//Skickar ett mail till i inställningar given mail med info om att någon har submittat en form. Är endast aktiverad när rutan i inställningar är ikryssad.
function sendFormSubmissionEmailToMe() {
  var settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inställningar");
  var settingsData = settings.getDataRange().getValues();

  if (settingsData[1][1] == true) {
    var myEmail = settingsData[2][1];
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Formulärsvar 1");
    var data = spreadsheet.getDataRange().getValues();
    var firstRow = data[0];
    var nameIdx = getColIdx(firstRow, "Namn på bokare (Kommer synas i bokningskalendern)");
    var startDateIdx = getColIdx(firstRow, "Bokningens starttid");
    var endDateIdx = getColIdx(firstRow, "Bokningens sluttid");
    var seatsIdx = getColIdx(firstRow, "Önskemål om antal säten i bilen (Kommer synas i bokningskalendern)");
    var foreningIdx = getColIdx(firstRow, "Förening, Kommitté, Funktionärspost (Kommer synas i bokningskalendern)");
    var targetIdx = getColIdx(firstRow, "Ändamål");
    var mobilIdx = getColIdx(firstRow, "Ditt telefonnummer (Kommer synas i bokningskalendern)");
    var emailIdx = getColIdx(firstRow, "Mailadress för bokningsbekräftelse");
    var errorIdx = getColIdx(firstRow, "Error messages");
    var errorMailIdx = getColIdx(firstRow, "Error mail sent");
    var confirmationMailIdx = getColIdx(firstRow, "Confirmation mail sent");

    //getColIdx returnerar -1 om strängen inte finns i någon kolumn på rad 1. Denna kodsnutt testar för felstavningar där.
    testArr = [nameIdx,startDateIdx,endDateIdx,seatsIdx,foreningIdx,mobilIdx,emailIdx,errorIdx,errorMailIdx,confirmationMailIdx, targetIdx];
    if (testArr.some(item => item == -1)) {
      throw new Error("Någon kolumnheader är felstavad. Dubbelkolla alla stavningar");
    }
    var lastRow = data.length - 1;
    var message = "Bokningsförfrågan av Helikoptern inkommen med följande formulärsvar:\nNamn på bokare: " + 
    data[lastRow][nameIdx] + "\nBokningens starttid: " + data[lastRow][startDateIdx] + "\nBokningens sluttid: " + data[lastRow][endDateIdx] + "\nFörening: " + data[lastRow][foreningIdx]
    + "\nÄndamål: " + data[lastRow][targetIdx] + ".";
    GmailApp.sendEmail(myEmail, "Bokningsförfrågan av Helikoptern", message);

  }




}

//Returnerar true om eventet [startDate,endDate] redan finns i eventCal
function checkForEventOverlap(eventCal, startDate, endDate) {
  var numberOfOverlapCases = eventCal.getEvents(startDate, endDate).length;
  return numberOfOverlapCases;
}

//Lägger till ett event till kalendern från datan i en given rad i kalkylarket.
function addEvent(eventCal, spreadsheet, row) {
  var data = spreadsheet.getDataRange().getValues();

  var firstRow = data[0];
  var nameIdx = getColIdx(firstRow, "Namn på bokare (Kommer synas i bokningskalendern)");
  var startDateIdx = getColIdx(firstRow, "Bokningens starttid");
  var endDateIdx = getColIdx(firstRow, "Bokningens sluttid");
  var seatsIdx = getColIdx(firstRow, "Önskemål om antal säten i bilen (Kommer synas i bokningskalendern)");
  var foreningIdx = getColIdx(firstRow, "Förening, Kommitté, Funktionärspost (Kommer synas i bokningskalendern)");
  var mobilIdx = getColIdx(firstRow, "Ditt telefonnummer (Kommer synas i bokningskalendern)");
  var emailIdx = getColIdx(firstRow, "Mailadress för bokningsbekräftelse");
  var errorIdx = getColIdx(firstRow, "Error messages");
  var errorMailIdx = getColIdx(firstRow, "Error mail sent");
  var confirmationMailIdx = getColIdx(firstRow, "Confirmation mail sent");

  //getColIdx returnerar -1 om strängen inte finns i någon kolumn på rad 1. Denna kodsnutt testar för felstavningar där.
  testArr = [nameIdx,startDateIdx,endDateIdx,seatsIdx,foreningIdx,mobilIdx,emailIdx,errorIdx,errorMailIdx,confirmationMailIdx];
  if (testArr.some(item => item == -1)) {
    throw new Error("Någon kolumnheader är felstavad. Dubbelkolla alla stavningar");
  }

  var startDate = new Date(data[row][startDateIdx]);
  var endDate = new Date(data[row][endDateIdx]);
  var email_address = data[row][emailIdx]

  //Sätter titel och beskrivning för kalendereventet, olika beroende på sektionsaktiv eller privatperson.
  if (data[row][foreningIdx] == "") { //Privatperson
    var title = data[row][nameIdx];
    var description = "<b>Antal säten:</b> " + data[row][seatsIdx] + "\n<b>Mobil:</b> " + data[row][mobilIdx];
  } else { //Sektionsaktiv
    var title = data[row][foreningIdx];
    var description = "<b>Antal säten:</b> " + data[row][seatsIdx] + "\n<b>Ansvarig:</b> " + data[row][nameIdx];
  }

  var cell_error = spreadsheet.getRange(row+1,errorIdx+1);
  var cell_error_mail = spreadsheet.getRange(row+1,errorMailIdx+1);
  var cell_confirmation_mail = spreadsheet.getRange(row+1,confirmationMailIdx+1);

  //Kollar om sluttid är före starttid. Om det är sant så skickar den ett mail till bokaren.
  if (startDate.getTime() > endDate.getTime()) {
    cell_error.setValue("Sluttid före starttid");
    cell_error.setFontColor("red");
    if (cell_error_mail.isBlank()) {
      try {
        sendFailedEmail(email_address, "Din bokning (" + startDate.getDate() + " " + monthName(startDate) + " " + startDate.getFullYear() + " " + String(startDate.getHours()).padStart(2, "0") + ":" + String(startDate.getMinutes()).padStart(2, "0") +  " till "
      + endDate.getDate() + " " + monthName(endDate) + " " + endDate.getFullYear() + " " + String(endDate.getHours()).padStart(2, "0") + ":" + String(endDate.getMinutes()).padStart(2, "0") + ") misslyckades tyvärr på grund av att du satte starttid senare än sluttid. Gör en ny bokning!");
        cell_confirmation_mail.setValue("Error mail sent");
        cell_confirmation_mail.setFontColor("green");
      } catch (mail_error) {
        cell_error_mail.setValue(mail_error);
        cell_error_mail.setFontColor("red");
      }
    }
    return;
  }

  //Kollar om eventet överlappar med tidigare bokningar. Om det är sant så skickar den ett mail till bokaren.
  if (checkForEventOverlap(eventCal, startDate, endDate)) {
    cell_error.setValue("Överlappar med andra bokningar");
    cell_error.setFontColor("red");
    if (cell_error_mail.isBlank()) {
      try {
        sendFailedEmail(email_address, "Din bokning (" + startDate.getDate() + " " + monthName(startDate) + " " + startDate.getFullYear() + " " + String(startDate.getHours()).padStart(2, "0") + ":" + String(startDate.getMinutes()).padStart(2, "0") +  " till "
      + endDate.getDate() + " " + monthName(endDate) + " " + endDate.getFullYear() + " " + String(endDate.getHours()).padStart(2, "0") + ":" + String(endDate.getMinutes()).padStart(2, "0") + ") misslyckades tyvärr på grund av att den delvis överlappar med en annan redan lagd bokning. Gör en ny bokning!");
        cell_confirmation_mail.setValue("Error mail sent");
        cell_confirmation_mail.setFontColor("green");
      } catch (mail_error) {
        cell_error_mail.setValue(mail_error);
        cell_error_mail.setFontColor("red");
      }
    }
    return;
  }

  //Försöker skapa eventet. Om det mislyckas får bokaren ett mail.
  try {
    eventCal.createEvent(title, startDate, endDate, {description: description});
  } catch (error) {
    console.error(error);
    cell_error.setValue(error);
    cell_error.setFontColor("red");
    if (cell_error_mail.isBlank()) {
      
      try {
        sendFailedEmail(email_address, "Din bokning (" + startDate.getDate() + " " + monthName(startDate) + " " + startDate.getFullYear() + " " + String(startDate.getHours()).padStart(2, "0") + ":" + String(startDate.getMinutes()).padStart(2, "0") +  " till "
      + endDate.getDate() + " " + monthName(endDate) + " " + endDate.getFullYear() + " " + String(endDate.getHours()).padStart(2, "0") + ":" + String(endDate.getMinutes()).padStart(2, "0") + ") misslyckades tyvärr på grund av följande fel: " + error + ". Gör en ny bokning!");
        cell_confirmation_mail.setValue("Skickade errormail!");
        cell_confirmation_mail.setFontColor("green");
      } catch (mail_error) {
        cell_error_mail.setValue(mail_error);
        cell_error_mail.setFontColor("red");
      }
    }
    return;
  }

  //Denna kod körs endast om eventet lyckades skapas (finns returns i varje error checking funktion ovan). Skickar även konfirmationsmail.
  cell_error.setValue("Lades till i kalendern!");
  cell_error.setFontColor("green");
  if (cell_confirmation_mail.isBlank()) {
    try {
    sendSuccessfulEmail(email_address, "Hej!\n\nDu har bokat Fysikteknologsektionens bil Helikoptern från "
      + startDate.getDate() + " " + monthName(startDate) + " " + startDate.getFullYear() + " " + String(startDate.getHours()).padStart(2, "0") + ":" + String(startDate.getMinutes()).padStart(2, "0") +  " till "
      + endDate.getDate() + " " + monthName(endDate) + " " + endDate.getFullYear() + " " + String(endDate.getHours()).padStart(2, "0") + ":" + String(endDate.getMinutes()).padStart(2, "0")
      + ".\nOm du bokade som privatperson måste du läsa igenom bokningsvillkoren och fylla i kontraktformuläret som finns på bilnisses sida på Ftek. Du är själv ansvarig att kontakta Bilnisse för att bestämma tid för nyckelutlämning.") 
    cell_confirmation_mail.setValue("Skickade bekräftelsemail!");
    cell_confirmation_mail.setFontColor("green");
    } catch (mail_error) {
      cell_confirmation_mail.setValue(mail_error);
      cell_confirmation_mail.setFontColor("red");
    }
  }

}

//Tar bort ALLA events från kalendern och gör om den från scratch från excelarket.
function remakeEntireCalendar() {
  var settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inställningar");
  var settingsData = settings.getDataRange().getValues();

  if (settingsData[0][1] == false) {

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Formulärsvar 1");
    var data = spreadsheet.getDataRange().getValues();
    var calendarId = "c_567s2i1r0slu9hmr6lmi53h5j8@group.calendar.google.com"
    var eventCal = CalendarApp.getCalendarById(calendarId);

    console.log("Tar bort alla events från kalendern...")
    var fromDate = new Date(2022,1,1,0,0,0);
    var toDate = new Date(2040,1,1,0,0,0);
    var events = eventCal.getEvents(fromDate, toDate);
    for(var i=0; i<events.length;i++) {
      events[i].deleteEvent();
    }

    console.log("Lägger till alla events från kalkylarket till kalendern...")
    for (row = 1; row < data.length; row++) {
      console.log("Rad: " + row);
      addEvent(eventCal, spreadsheet, row);
    }
  }
}

//Lägger till eventet längst ned i excelarket (senaste inskickade formulärsvaret alltså).
function addLatestEvent() {
  var settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inställningar");
  var settingsData = settings.getDataRange().getValues();

  if (settingsData[0][1] == false) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Formulärsvar 1");
    var data = spreadsheet.getDataRange().getValues();
    var calendarId = "c_567s2i1r0slu9hmr6lmi53h5j8@group.calendar.google.com"
    var eventCal = CalendarApp.getCalendarById(calendarId);

    addEvent(eventCal, spreadsheet, data.length-1);
  }
}


//Körs när excelarket öppnas manuellt. Skapar en knapp i menyraden med lite nice funktioner :)
function onOpen() {
  var ui = SpreadsheetApp.getUi()
  ui.createMenu("Lite funktioner")
  .addItem("Återskapa hela kalendern","remakeEntireCalendar")
  //.addItem("Skicka mail","sendMail")
  .addToUi();

  //var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Formulärsvar 1");
  //var cell = spreadsheet.getRange(6,5);
  //cell.setValue(1);
}

//Funktion där man manuellt kan lägga till/ta bort bokningar. Körs när excelarket editas. Måste aktiveras i inställningar.
function editor(ee) {

  if (SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName() == "Formulärsvar 1") {

    

    var settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inställningar");
    var settingsData = settings.getDataRange().getValues();

    if (settingsData[0][1] == true) {

      const range = ee.range;
      var row = range.getRow();
      var column = range.getColumn();
      var value = range.getCell(1, 1).getValue();

      if (column == 1) {

        SpreadsheetApp.getActive().toast("Gör inga ändringar i dokumentet", "Laddar...");

        var calendarId = "c_567s2i1r0slu9hmr6lmi53h5j8@group.calendar.google.com"
        var eventCal = CalendarApp.getCalendarById(calendarId);
        var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Formulärsvar 1");
        var data = spreadsheet.getDataRange().getDisplayValues();

        var firstRow = data[0];
        var startDateIdx = getColIdx(firstRow, "Bokningens starttid");
        var endDateIdx = getColIdx(firstRow, "Bokningens sluttid");
        var nameIdx = getColIdx(firstRow, "Namn på bokare (Kommer synas i bokningskalendern)");
        var foreningIdx = getColIdx(firstRow, "Förening, Kommitté, Funktionärspost (Kommer synas i bokningskalendern)");

        if (value == "") {
          var fromDateString = String(data[row - 1][startDateIdx]);
          var toDateString = String(data[row - 1][endDateIdx]);
          fixedFromDateString = fromDateString.replace(/\./g, ":");
          fixedToDateString = toDateString.replace(/\./g, ":");
          var fromDate = new Date(Date.parse(fixedFromDateString));
          var toDate = new Date(Date.parse(fixedToDateString));

          var events = eventCal.getEvents(fromDate, toDate);
          var eventDeleted = false;
          var eventIndex = 0;

          for (let i = 0; i < events.length; i++) {
            if (events[i].getTitle() == data[row - 1][nameIdx] || events[i].getTitle() == data[row - 1][foreningIdx]) {
              events[i].deleteEvent();
              eventIndex = i;
              eventDeleted = true;
            }
          }
          if (eventDeleted) {
            if (data[row - 1][foreningIdx]) {
            SpreadsheetApp.getActive().toast("Tog bort kalenderhändelse för " + data[row - 1][foreningIdx] + ", bokad av " + data[row - 1][nameIdx], "Kalenderhändelse raderad");
            } else {
              SpreadsheetApp.getActive().toast("Tog bort kalenderhändelse bokad av " + data[row - 1][nameIdx] + " som privatperson.", "Kalenderhändelse raderad");
            }
          } else {
            SpreadsheetApp.getActive().toast("", "Error");
            SpreadsheetApp.getUi().alert("Kunde inte hitta något event i kalendern som matchar datan på given rad. Dubbelkolla excelarkets data med kalendern.");
            throw new Error("Kunde inte hitta något event i kalendern som matchar datan på given rad. Dubbelkolla excelarkets data med kalendern.");
          }

        } else {
          addEvent(eventCal, spreadsheet, row-1);
          if (data[row - 1][foreningIdx]) {
            SpreadsheetApp.getActive().toast("Lade till kalenderhändelse för " + data[row - 1][foreningIdx] + ", bokad av " + data[row - 1][nameIdx], "Lade till kalenderhändelse");
            } else {
              SpreadsheetApp.getActive().toast("Lade till kalenderhändelse bokad av " + data[row - 1][nameIdx] + " som privatperson.", "Lade till kalenderhändelse");
            }
        }
      }

    }
  }
}

//Körs när datan i någon cell i excelarket ändras manuellt.
//function onEdit(e) {
//  editor(e);
//}


//Körs när den highlightade cellen i excelarket ändras
//function onSelectionChange() {
//  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Formulärsvar 1");
  //var cell = spreadsheet.getRange(5,5);
  //cell.setValue(3);
//}
