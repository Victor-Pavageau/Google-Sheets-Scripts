const getCellValue = (sheet, cell) => sheet.getRange(cell).getValue();
const setCellValue = (sheet, cell, value) =>
  sheet.getRange(cell).setValue(value);
const getSheet = (sheetName) =>
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

/* 
---------------------
SECTION 1: WEEK COUNT - Input Number
---------------------
*/

const GLOBAL_SHEET = getSheet('Global planning');
const WEEK_COUNT_CELL = 'B39';
const OPERATION_VALUE_CELL = 'G38:G39';

const AddWeeks = () => {
  setCellValue(
    GLOBAL_SHEET,
    WEEK_COUNT_CELL,
    getOldWeekCount() + getNumberOfWeeksToAddOrRemove(),
  );
  resetOperationValue();
};

const RemoveWeeks = () => {
  setCellValue(
    GLOBAL_SHEET,
    WEEK_COUNT_CELL,
    getOldWeekCount() - getNumberOfWeeksToAddOrRemove(),
  );
  resetOperationValue();
};

const getOldWeekCount = () => {
  return getCellValue(GLOBAL_SHEET, WEEK_COUNT_CELL);
};

const getNumberOfWeeksToAddOrRemove = () => {
  return getCellValue(GLOBAL_SHEET, OPERATION_VALUE_CELL);
};

const resetOperationValue = () => {
  setCellValue(GLOBAL_SHEET, OPERATION_VALUE_CELL, 0);
};

/* 
---------------------
SECTION 1: RANDOM MOVIE - Get Random Movie From List
---------------------
*/

const MOVIE_SHEET = getSheet('Movies');
const START_INDEX = 6;
const COUNT_MOVIES_CELL = 'C5';
const COLUMNS = ['B', 'C'];

const selectRandomMovie = () => {
  const cellCoordinates = getRandomMovieCellCoordinates();
  MOVIE_SHEET.setActiveSelection(MOVIE_SHEET.getRange(cellCoordinates));
  updateMovieCount();
};

const getRandomMovieCellCoordinates = () => {
  const randomColumn = COLUMNS[Math.floor(Math.random() * COLUMNS.length)];
  const randomIndex =
    Math.floor(
      Math.random() * (getLastMovieIndex(randomColumn) + 1 - START_INDEX),
    ) + START_INDEX;

  return `${randomColumn}${randomIndex}`;
};

const getLastMovieIndex = (col) => {
  for (
    var index = START_INDEX;
    getCellValue(MOVIE_SHEET, `${col}${index}`).length > 1;
    index++
  );

  return index - 1;
};

const getRemainingMoviesCount = () => {
  let remainingMoviesCount = 0;
  COLUMNS.forEach((column) => {
    remainingMoviesCount += getLastMovieIndex(column) + 1 - START_INDEX;
  });

  return remainingMoviesCount;
};

const updateMovieCount = () => {
  setCellValue(
    MOVIE_SHEET,
    COUNT_MOVIES_CELL,
    `Remaining movies : ${getRemainingMoviesCount()}`,
  );
};
