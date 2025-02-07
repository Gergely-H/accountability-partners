const addTimezoneSelectorConfig = () => {
  const sheet = getHabitSheet();
  const columnIndex = convertColumnNameToIndex(CONFIG_COLUMN_NAME);
  const configLabelRowIndex = findConfigLabelRowIndex(TIMEZONE_CONFIG_LABEL);
  const configLabelCell = sheet.getRange(configLabelRowIndex, columnIndex);
  const configCell = getCellRelativeToCell(configLabelCell, 1, 0, sheet);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(UNIQUE_UTC_TIMEZONES, true)
    .build();

  configCell.setDataValidation(rule);

  configCell.setValue(DEFAULT_TIMEZONE);
};

const initializeSheetAndTriggers = () => {
  addTimezoneSelectorConfig();
  updateDailyTrigger();
  createOnEditTriggerForTimezoneConfig();
  setDate();
};
