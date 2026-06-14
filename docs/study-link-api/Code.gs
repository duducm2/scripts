/**
 * Study Link API — Google Apps Script for sheet "my-study-database"
 *
 * Sheet layout (active tab):
 *   A1 — YouTube  (key subtopic)
 *   A2 — Article  (key subtopic_article)
 *   A3 — Favorite (key subtopic_favorite)
 *
 * Deploy: Deploy → Manage deployments → Edit → New version → Deploy
 * Web app: Execute as Me, Who has access: Anyone
 *
 * GET ?key=<key>           → raw stored body (PC / AHK)
 * GET ?key=<key>&open=1    → HTML redirect to URL after url= (MacroDroid Get macros)
 * POST body key=<key>&url= → stored verbatim in the matching cell
 */

function doPost(e) {
  var key = e.parameter && e.parameter.key ? String(e.parameter.key) : "";
  var contents = e.postData && e.postData.contents ? e.postData.contents : "";

  if (key === "subtopic_article" || contents.indexOf("subtopic_article") !== -1)
    return doPostArticle(e);
  if (
    key === "subtopic_favorite" ||
    contents.indexOf("subtopic_favorite") !== -1
  )
    return doPostFavorite(e);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("A1").setValue(contents);
  return ContentService.createTextOutput("Saved");
}

function doGet(e) {
  var key = e.parameter && e.parameter.key ? String(e.parameter.key) : "";

  if (key === "subtopic_article") return doGetArticle(e);
  if (key === "subtopic_favorite") return doGetFavorite(e);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var link = sheet.getRange("A1").getValue();
  if (e.parameter && e.parameter.open === "1") {
    return openStoredLinkResponse(
      link,
      "No video link stored. Run Set Video on your phone first.",
    );
  }
  return ContentService.createTextOutput(link);
}

function doPostArticle(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("A2").setValue(e.postData.contents);
  return ContentService.createTextOutput("Saved");
}

function doGetArticle(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var link = sheet.getRange("A2").getValue();
  if (e.parameter && e.parameter.open === "1") {
    return openStoredLinkResponse(
      link,
      "No article link stored. Run Set Article on your phone first.",
    );
  }
  return ContentService.createTextOutput(link);
}

function doPostFavorite(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("A3").setValue(e.postData.contents);
  return ContentService.createTextOutput("Saved");
}

function doGetFavorite(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var link = sheet.getRange("A3").getValue();
  if (e.parameter && e.parameter.open === "1") {
    return openStoredLinkResponse(
      link,
      "No favorite link stored. Run Set Favorite on your phone first.",
    );
  }
  return ContentService.createTextOutput(link);
}

function extractUrlFromStoredBody(stored) {
  var s = String(stored || "");
  var pos = s.indexOf("url=");
  if (pos !== -1) return s.substring(pos + 4);
  return s;
}

function openStoredLinkResponse(stored, emptyMessage) {
  var url = extractUrlFromStoredBody(stored);
  if (!url || url.indexOf("http") !== 0) {
    return HtmlService.createHtmlOutput("<p>" + emptyMessage + "</p>");
  }
  return HtmlService.createHtmlOutput(
    "<script>window.location.href=" + JSON.stringify(url) + ";</script>",
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
