const setDate = () => {
  const newDate = new Date();
  const dd = String(newDate.getDate()).padStart(2, '0');
  const mm = String(newDate.getMonth() + 1).padStart(2, '0');
  const yyyy = newDate.getFullYear();

  const today = `${yyyy}/${mm}/${dd}`;

  const sheet = getHabitSheet();
  const dateColumnIndex = convertColumnNameToIndex(FIRST_DATE_COLUMN_NAME);
  const dateCell = sheet.getRange(DATE_ROW_INDEX, dateColumnIndex);
  dateCell.setValue(today);
};

/**
 * This function runs daily.
 * Its name must be equal to the NEW_DAY_HANDLER_FUNCTION_NAME constant.
 */
const insertNewDayColumn = () => {
  const dateColumnIndex = convertColumnNameToIndex(FIRST_DATE_COLUMN_NAME);
  const sheet = getHabitSheet();
  sheet.insertColumnBefore(dateColumnIndex);

  setDate();
};
