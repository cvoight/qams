/* ****************************************************************************
 * This Google Apps Script is used to track and manage set production in a
 * Google Sheets answer spreadsheet. It uses writer tags and the background
 * color of cells to determine how many questions a writer has claimed, and
 * then uses the bolded status of a cell to determine whether a question has
 * been finished (edited and placed in a packet, or written depending on how
 * you instruct your writers and editors). It also tracks subcategory
 * completion. The template spreadsheet file that accompanies this script can
 * be found at:
 * https://docs.google.com/spreadsheets/d/16tlqgO5As9mKEj1t-89SmVSMg_vRPW-CliCMTA7TQTE/edit#gid=564996451
 * Author: Cody Voight
 * Version: 0.9.9-gamma1
 * ***************************************************************************/

function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu("qams")
      .addItem("Writer completion", "writerCompletion")
      .addItem("Subcategory completion", "subcategoryCompletion")
      .addToUi();
}

// qams uses named ranges to collect and set data.

// the answer space, structured with packets as column headers and tossups
// and bonuses in alternating rows (with tossups as Row 1 of the range).
// if your answer space is structured differently, see the comments in the
// writerCompletion function below.
var answerRange = "answers";
// 1-column writer tag range, with writer background color
// a common writer tag is a writer's initials in square brackets
var tagRange = "tags";
// 1-column ranges to set claimed tossups and bonuses, by writer
var claimedTossupsRange = "claimedTossups";
var claimedBonusesRange = "claimedBonuses";
// 1-column ranges to set finished tossups and bonuses, by writer
var finishedTossupsRange = "finishedTossups";
var finishedBonusesRange = "finishedBonuses";
// 1-column range to set finished questions, by subcategory and question type
var subcategoryCompletionRange = "subcategoryCompletion";

function writerCompletion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var answerSpace = ss.getRangeByName(answerRange);
  var writerTags = ss.getRangeByName(tagRange);
  var claimedTossups = ss.getRangeByName(claimedTossupsRange);
  var claimedBonuses = ss.getRangeByName(claimedBonusesRange);
  var finishedTossups = ss.getRangeByName(finishedTossupsRange);
  var finishedBonuses = ss.getRangeByName(finishedBonusesRange);
  var answerSpaceBackgrounds = answerSpace.getBackgrounds();
  var answerSpaceFontWeights = answerSpace.getFontWeights();
  var answerSpaceValues = answerSpace.getValues();
  var writerTagColors = writerTags.getBackgrounds();
  var writerTagValues = writerTags.getValues();
  var claimedTossupsTotal = 0, claimedBonusesTotal = 0;
  var finishedTossupsTotal = 0, finishedBonusesTotal = 0;

  for (var m in writerTagColors) {
    for (var p in answerSpaceBackgrounds) {
      for (var q in answerSpaceBackgrounds[p]) {
        if (writerTagValues[m][0].toString().length > 0) {
          if (answerSpaceValues[p][q].toString().indexOf(
                  writerTagValues[m][0].toString()) > -1) {
            answerSpaceBackgrounds[p][q] = writerTagColors[m][0];
          }
        }
        if (answerSpaceBackgrounds[p][q] === writerTagColors[m][0]) {
          if (p % 2 === 0 || p === 0) {
            claimedTossupsTotal += 1;
          } else {
            claimedBonusesTotal += 1;
          }
          if (answerSpaceValues[p][q].toString().length === 0) {
            continue;
          }
          if (answerSpaceFontWeights[p][q] === "bold") {
            // Replace 'p' with 'q' if you have tossups and bonuses in
            // alternating columns instead of alternating rows.
            if (p % 2 === 0 || p === 0) {
              // Switch finishedTossupsTotal and finishedBonusesTotal if
              // bonuses start the alternating sequence, instead of tossups.
              finishedTossupsTotal += 1;
            } else {
              finishedBonusesTotal += 1;
            }
          }
        }
      }
    }
    claimedTossups.getCell(+m + 1, 1).setValue(claimedTossupsTotal);
    claimedBonuses.getCell(+m + 1, 1).setValue(claimedBonusesTotal);
    finishedTossups.getCell(+m + 1, 1).setValue(finishedTossupsTotal);
    finishedBonuses.getCell(+m + 1, 1).setValue(finishedBonusesTotal);
    claimedTossupsTotal = claimedBonusesTotal = 0;
    finishedTossupsTotal = finishedBonusesTotal = 0;
  }
  answerSpace.setBackgrounds(answerSpaceBackgrounds);
}

function subcategoryCompletion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var answerSpace = ss.getRangeByName(answerRange);
  var answerSpaceFontWeights = answerSpace.getFontWeights();
  var answerSpaceValues = answerSpace.getValues();
  var subcategory = ss.getRangeByName(subcategoryCompletionRange);
  var subcategoryTotal = 0;

  for (var p in answerSpaceFontWeights) {
    for (var q in answerSpaceFontWeights[p]) {
      if (answerSpaceFontWeights[p][q] === "bold" &&
          answerSpaceValues[p][q].toString().length > 0) {
        subcategoryTotal += 1;
      }
    }
    subcategory.getCell(+p + 1, 1).setValue(subcategoryTotal);
    subcategoryTotal = 0;
  }
}
