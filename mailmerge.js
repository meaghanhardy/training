function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [ {name: "Send Draft Email", functionName: "sendEmails"}];
  ss.addMenu("Weekly Email", menuEntries);
};




  function sendEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var countsSheet = ss.getSheets()[0];
  var dataRange = countsSheet.getRange("A2:B9").getValues();
  var data = dataRange.join('<p><li>');
  var email = countsSheet.getRange("G2")

  var ListsSheet = ss.getSheets()[1];
 
  var EAClientUpdates_raw = ListsSheet.getRange("A2:A").getValues(); 
  var EAClientUpdates_input = EAClientUpdates_raw.filter(String);
  var EAClientUpdates = EAClientUpdates_input.join('<p style="padding-left: 30px;"><li>');

  var NGPClientUpdates_raw = ListsSheet.getRange("B2:B").getValues();
  var NGPClientUpdates_input = NGPClientUpdates_raw.filter(String);
  var NGPClientUpdates = NGPClientUpdates_input.join('<p style="padding-left: 30px;"><li>');

  var EA8Migrations_raw = ListsSheet.getRange("C2:C").getValues();
  var EA8Migrations_input = EA8Migrations_raw.filter(String);
  var EA8Migrations = EA8Migrations_input.join('<p style="padding-left: 30px;"><li>');


  var NGPDemosOffered_raw = ListsSheet.getRange("D4:D").getValues();
  var NGPDemosOffered_input = NGPDemosOffered_raw.filter(String);
  var NGPDemosOffered = NGPDemosOffered_input.join('<p style="padding-left: 30px;"><li>');

  var NGPDemosScheduled_raw = ListsSheet.getRange("G4:G").getValues();
  var NGPDemosScheduled_input = NGPDemosScheduled_raw.filter(String);
  var NGPDemosScheduled = NGPDemosScheduled_input.join('<p style="padding-left: 30px;"><li>');

  var NGPMigrationsScheduled_raw = ListsSheet.getRange(3, 10, ListsSheet.getMaxRows() - 1, 1).getValues();
  var NGPMigrationsScheduled_input = NGPMigrationsScheduled_raw.filter(String);
  var NGPMigrationsScheduled = NGPMigrationsScheduled_input.join('<p style="padding-left: 30px;"><li>');


  var bodya = "<p>Hey all!</p><p>Sending around the queue information for this week.&nbsp;</p><p>&nbsp;</p><p><strong>Conversions&nbsp;</strong>(new clients to NGP/EA)</p><p>Calendar when you click&nbsp;this link</p><p>&nbsp;</p><p><em>Relevant counts:</em></p><li>"  
  var bodyb = "<p><strong>EA Enterprise/AM client updates:</strong></p><li>"
  var bodyc = "<p>&nbsp;</p><p><strong>NGP Enterprise/AM client updates:</strong></p><li>"
  var bodyd = "<p><strong>Migrations (</strong>Existing NGP/EA clients moving to a new platform<strong>)</strong></p><p>&nbsp;</p><p><strong>NGP Migrations:</strong></p><p><span style='text-decoration: underline;'>Demos Offered</span></p><li>"
  var bodye = "<p><span style='text-decoration: underline;'>Demos scheduled:</span></p><li>"
  var bodyf = "<p><span style='text-decoration: underline;'>Migrations scheduled:</span></p><li>"
  var bodyg = "<p>&nbsp;</p><p><strong>EA8 Migrations</strong>: note - none of these are currently scheduled, but they are the top tier to move</p><li>"

  var emailSubject = "[DRAFT] Migrations and Conversions Queue";

  var fullemail = bodya + data  + bodyb + EAClientUpdates + bodyc + NGPClientUpdates + bodyd + NGPDemosOffered+ bodye + NGPDemosScheduled + bodyf + NGPMigrationsScheduled + bodyg + EA8Migrations

        MailApp.sendEmail({
          to:"meaghan.e.hardy@gmail.com", 
          subject: emailSubject,
          htmlBody: fullemail
                      });
 }

