/* ****************************************************************************
 * This Google Apps Script is used to track and manage set production in a
 * Google Sheets answer spreadsheet.
 * Writer completion: uses writer tags and the background color of cells to
 * determine how many questions a writer has claimed, and then uses the bolded
 * status of a cell to determine whether a question has been finished (edited
 * and placed in a packet, or written depending on how you instruct your writers
 * and editors).
 * Subcategory completion: determines how many questions in each subcategory
 * have been completed.
 * Packet templates: generates packet templates based on a user-defined
 * distribution specification.
 * https://docs.google.com/spreadsheets/d/16tlqgO5As9mKEj1t-89SmVSMg_vRPW-CliCMTA7TQTE/edit#gid=564996451
 * Author: Cody Voight
 * Version: 0.9.9-gamma.2
 * ***************************************************************************/

/**
 * @OnlyCurrentDoc
 */

// qams uses named ranges to collect and set data. It does not write or read
// data outside the ranges.

// the answer space, structured with packets as column headers and tossups
// and bonuses in alternating rows (with tossups as Row 1 of the range).
// if your answer space is structured differently, see the comments in the
// writerCompletion function below.
const answerRangeString = "answers";
// 1-column writer tag range, with writer background color
// a common writer tag is a writer's initials in square brackets
// writer tags must be a unique string
const tagRangeString = "tags";
// 1-column ranges to set claimed tossups and bonuses, by writer
const claimedTossupsRangeString = "claimedTossups";
const claimedBonusesRangeString = "claimedBonuses";
// 1-column ranges to set finished tossups and bonuses, by writer
const finishedTossupsRangeString = "finishedTossups";
const finishedBonusesRangeString = "finishedBonuses";
// 1-column range to set finished questions, by subcategory and question type
const subcategoryCompletionRangeString = "subcategoryCompletion";
// 1-column range to get distribution
const distributionRangeString = "distribution";
// the template space, structured with packets as rows (alternating tossups and
// bonuses) and question indices as columns
const templatesRangeString = "templates";

/******************************************************************************/

// creates the qams menu
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("qams")
    .addItem("Writer completion", "writerCompletion")
    .addItem("Subcategory completion", "subcategoryCompletion")
    .addSeparator()
    .addItem("Packet templates", "packetTemplates")
    .addToUi();
}

/******************************************************************************/

// global helper functions

// zips 3 arrays by index with a function applied to them
const zip3With = (f, a, b, c) => a.map((e, i) => f(e, b[i], c[i]));
// returns the index of a string in an array if contained in the input string,
// or -1 otherwise. does not match against the empty string.
const indexIfIncludes = (a, s) =>
  a.findIndex((e) => !!e.length && s.includes(e));
// splits an array into two arrays. the first array contains elements whose
// indices returned true for f. the second array contains elements whose
// indices returned false for f.
const partitionByIndex = (f, a) =>
  a.reduce((result, e, i) => (result[f(i) ? 0 : 1].push(e), result), [[], []]);
// chunks an array into an array of arrays with length d
const chunk = (a, d) =>
  a.reduce((r, e, i) => {
    const c = Math.floor(i / d);
    !r[c] ? r.push([e]) : r[c].push(e);
    return r;
  }, []);
// chunks a Map's values to a column array
const chunkMapToColumn = (v) => Array.from(v, (e) => [e]);

/******************************************************************************/

function writerCompletion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // getRangeByName returns a named range, or null if not found
  const answerSpaceRange = ss.getRangeByName(answerRangeString);
  const tagsRange = ss.getRangeByName(tagRangeString);
  const claimedTossupsRange = ss.getRangeByName(claimedTossupsRangeString);
  const claimedBonusesRange = ss.getRangeByName(claimedBonusesRangeString);
  const finishedTossupsRange = ss.getRangeByName(finishedTossupsRangeString);
  const finishedBonusesRange = ss.getRangeByName(finishedBonusesRangeString);

  const errorAlertTitle = "qams execution error";
  const namedRangeNotFound =
    "A named range could not be found. Please\
      double-check the named ranges against the script variables.";
  const columnRangesUnequal =
    "The 1-column named ranges do not have the same\
      number of rows. Please double-check them.";
  /**
   * @param {SpreadsheetApp.Range[]} ranges
   */
  const rangeVerification = function (...ranges) {
    if (ranges.some((e) => e === null)) {
      throw new Error(namedRangeNotFound);
    } else if (
      ranges
        .filter((e) => e.getWidth() === 1)
        .map((e) => e.getHeight())
        .some((e, _, a) => e !== a[0])
    ) {
      throw new Error(columnRangesUnequal);
    } else {
      return true;
    }
  };
  try {
    rangeVerification(
      answerSpaceRange,
      tagsRange,
      claimedTossupsRange,
      claimedBonusesRange,
      finishedTossupsRange,
      finishedBonusesRange
    );
  } catch (e) {
    ui.alert(errorAlertTitle, e, ui.ButtonSet.OK);
    return;
  }

  // from Google Apps Script Developer Reference. return values dictate the
  // error checking required (minimal, e.g. no null checking is required).

  // getBackgrounds returns String[][], a two-dimensional array of color codes.
  // Color codes are in CSS notation, such as '#ffffff' or 'white'. A null value
  // resets the color.
  const answerSpaceBackgrounds = answerSpaceRange.getBackgrounds();
  // getValues returns Object[][], a two-dimensional array of values. Values
  // have type Number, Boolean, Date, or String. Empty cells are represented by
  // an empty string.
  const answerSpaceValues = answerSpaceRange.getValues();
  // getFontWeights returns String[][], a two-dimensional array of font weights.
  // Font weight is either 'bold' or 'normal'. A null value resets the font
  // weight.
  const answerSpaceFontWeights = answerSpaceRange.getFontWeights();
  const tagColors = tagsRange.getBackgrounds();
  const tagValues = tagsRange.getValues();
  const tagColorsFlat = tagColors.flat();
  const tagValuesFlat = tagValues.flat();

  const divisor = answerSpaceBackgrounds[0].length;

  // sets background based on writer tag & resets font weight to normal for
  // empty cells
  const correctBackgroundsValuesFontWeights = (b, v, f) =>
    !v.length
      ? [b, "normal"]
      : [tagColorsFlat[indexIfIncludes(tagValuesFlat, v)] || b, f];
  // assumes the first row of the answer space is tossups
  const isTossup = (i) => i % (divisor * 2) < divisor;
  // reducers
  const claimedReducer = (m, e) =>
    m.has(e[0]) ? m.set(e[0], (m.get(e[0]) || 0) + 1) : m;
  const finishedReducer = (m, e) =>
    m.has(e[0]) && e[1] === "bold" ? m.set(e[0], (m.get(e[0]) || 0) + 1) : m;

  // correct the answer space for further processing
  const correctedAnswerSpace = zip3With(
    correctBackgroundsValuesFontWeights,
    answerSpaceBackgrounds.flat(),
    answerSpaceValues.flat(),
    answerSpaceFontWeights.flat()
  );
  const [tossups, bonuses] = partitionByIndex(isTossup, correctedAnswerSpace);
  // Map(tagColors) preserves order of writers
  const claimedTossups = tossups.reduce(claimedReducer, new Map(tagColors));
  const claimedBonuses = bonuses.reduce(claimedReducer, new Map(tagColors));
  const finishedTossups = tossups.reduce(finishedReducer, new Map(tagColors));
  const finishedBonuses = bonuses.reduce(finishedReducer, new Map(tagColors));

  const [correctedBackgrounds, correctedFontWeights] =
    correctedAnswerSpace.reduce(
      (r, e, i) => (r[0].push(e[0]), r[1].push(e[1]), r),
      [[], []]
    );

  const columnRangeSetValueUnequal =
    "Cannot set writer claimed & finished\
      values. Most likely, a writer color has been re-used.";
  answerSpaceRange.setBackgrounds(chunk(correctedBackgrounds, divisor));
  answerSpaceRange.setFontWeights(chunk(correctedFontWeights, divisor));
  try {
    claimedTossupsRange.setValues(chunkMapToColumn(claimedTossups.values()));
    claimedBonusesRange.setValues(chunkMapToColumn(claimedBonuses.values()));
    finishedTossupsRange.setValues(chunkMapToColumn(finishedTossups.values()));
    finishedBonusesRange.setValues(chunkMapToColumn(finishedBonuses.values()));
  } catch (e) {
    ui.alert(
      errorAlertTitle,
      new Error(columnRangeSetValueUnequal),
      ui.ButtonSet.OK
    );
  }
}

/******************************************************************************/

function subcategoryCompletion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const answerSpaceRange = ss.getRangeByName(answerRangeString);
  const subcategoryRange = ss.getRangeByName(subcategoryCompletionRangeString);

  const errorAlertTitle = "qams execution error";
  const namedRangeNotFound =
    "A named range could not be found. Please\
      double-check the named ranges against the script variables.";
  const rowsUnequal =
    "The answer space & subcategory completion named ranges\
      do not have the same number of rows. Please double-check them.";
  /**
   * @param {SpreadsheetApp.Range[]} ranges
   */
  const rangeVerification = function (...ranges) {
    if (ranges.some((e) => e === null)) {
      throw new Error(namedRangeNotFound);
    } else if (ranges.map((e) => e.getHeight()).some((e, _, a) => e !== a[0])) {
      throw new Error(rowsUnequal);
    } else {
      return true;
    }
  };
  try {
    rangeVerification(answerSpaceRange, subcategoryRange);
  } catch (e) {
    ui.alert(errorAlertTitle, e, ui.ButtonSet.OK);
    return;
  }

  // getFontWeights returns String[][], a two-dimensional array of font weights.
  // Font weight is either 'bold' or 'normal'. A null value resets the font
  // weight.
  const answerSpaceFontWeights = answerSpaceRange.getFontWeights();
  const finishedSubcategories = answerSpaceFontWeights.map((e) =>
    e.reduce((r, e) => (e === "bold" ? [r[0] + 1] : r), [0])
  );
  subcategoryRange.setValues(finishedSubcategories);
}

/******************************************************************************/
function packetTemplates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const distributionRange = ss.getRangeByName(distributionRangeString);
  const templatesRange = ss.getRangeByName(templatesRangeString);

  const frequencyMap2 = (m, e) =>
    m.set(e.substring(0, 2), (m.get(e.substring(0, 2)) || 0) + 1);
  const frequencyMap4 = (m, e) =>
    m.set(e.substring(0, 4), (m.get(e.substring(0, 4)) || 0) + 1);
  const shuffle = (r, e, i) => {
    const rand = Math.floor(Math.random() * (i + 1));
    r[i] = r[rand];
    r[rand] = e;
    return r;
  };
  const swap = (xs, a, b) => (([xs[a], xs[b]] = [xs[b], xs[a]]), xs);

  const csp = (constraints) => {
    const constrain = (variables) => {
      return constraints.reduce((p, c) => {
        const [h, t, b, n, constraint] = c;
        return constraint(variables.slice(h, t + 1), b, n) ? p + 1 : p;
      }, 0);
    };

    const solve = (variables, level = 0) => {
      if (level > 5) return solve(variables.reduce(shuffle, []));
      if (constrain(variables) === 0) return variables;

      const indices = Array.from(
        { length: variables.length },
        (_, i) => i
      ).flatMap((i1, i, a) => a.slice(i + 1).map((i2) => [i1, i2]));
      const swaps = indices.map((e) => swap(variables.slice(), e[0], e[1]));
      const mc = swaps.reduce((p, c) => (constrain(c) < constrain(p) ? c : p));

      return solve(mc, level + 1);
    };

    return solve;
  };

  const packetizingProblem = (distribution) => {
    const size = distribution.length;
    const variables = distribution.reduce(shuffle, []);

    const consecutive = Array.from({ length: size - 1 }, (_, i) => [
      i,
      i + 1,
      "",
      "",
      (v, b, n) => v[0].substring(0, 2) === v[1].substring(0, 2),
    ]);

    const categoryFrequency = variables.reduce(frequencyMap2, new Map());
    const categories = Array.from(categoryFrequency.entries(), ([k, v]) => {
      if (v < 2) return [];
      const code = k.substring(1, 2);
      if (code !== "A" && code !== "B") return [];
      const width = code === "A" ? size / v : size / 2;
      const n = code === "A" ? 1 : Math.ceil(v / 2);
      return Array.from({ length: code === "B" ? 2 : v }, (_, i) => [
        Math.round(i * width),
        Math.round((i + 1) * width - 1),
        k,
        n,
        (v, b, n) =>
          v.reduce((p, c) => (c.substring(0, 2) === b ? p + 1 : p), 0) > n,
      ]);
    }).flat();

    const subcategoryFrequency = variables.reduce(frequencyMap4, new Map());
    const subcategories = Array.from(
      subcategoryFrequency.entries(),
      ([k, v]) => {
        if (v < 2) return [];
        const code = k.substring(3, 4);
        if (code !== "a" && code !== "b") return [];
        const width = code === "a" ? size / v : size / 2;
        const n = code === "a" ? 1 : Math.ceil(v / 2);
        return Array.from({ length: code === "b" ? 2 : v }, (_, i) => [
          Math.round(i * width),
          Math.round((i + 1) * width - 1),
          k,
          n,
          (v, b, n) =>
            v.reduce((p, c) => (c.substring(0, 4) === b ? p + 1 : p), 0) > n,
        ]);
      }
    ).flat();

    const constraints = consecutive.concat(categories, subcategories);
    return { variables, constraints };
  };

  const distribution = distributionRange
    .getValues()
    .flatMap((e) => (e[0] === "" ? [] : e));

  const templates = templatesRange.getValues().map((e) => {
    if (e[0] !== "") {
      const { variables, constraints } = packetizingProblem(distribution);
      const solve = csp(constraints);
      const template = solve(variables);
      e.splice(2, template.length, ...template);
    }
    return e;
  });

  templatesRange.setValues(templates);
}
