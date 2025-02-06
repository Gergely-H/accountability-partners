const findConfigLabelRowIndex = (configLabel, sheet = getHabitSheet()) => {
  const lastRow = sheet.getLastRow();
  const configColumnIndex = convertColumnNameToIndex(CONFIG_COLUMN_NAME);
  const column = sheet.getRange(FIRST_ACCOUNTABILITY_PARTNER_ROW_INDEX, configColumnIndex, lastRow, 1);
  const values = column.getValues();

  let relativeConfigRowIndex = 0;
  while (relativeConfigRowIndex < values.length && !values[relativeConfigRowIndex][0].toString().includes(configLabel)) {
    relativeConfigRowIndex++;
  }

  const absoluteConfigRowIndex = relativeConfigRowIndex + FIRST_ACCOUNTABILITY_PARTNER_ROW_INDEX;

  return absoluteConfigRowIndex;
};

const findConfigLabelOccurrenceCount = (configLabel, sheet = getHabitSheet()) => {
  const lastRow = sheet.getLastRow();
  const configColumnIndex = convertColumnNameToIndex(CONFIG_COLUMN_NAME);
  const column = sheet.getRange(FIRST_ACCOUNTABILITY_PARTNER_ROW_INDEX, configColumnIndex, lastRow, 1);
  const values = column.getValues().flat();

  const configLabelOccurenceCount = values.filter((cellValue) => cellValue.toString().includes(configLabel)).length;
  
  return configLabelOccurenceCount;
};

const handleAccountabilityPartnerCountConfig = (sheet = getHabitSheet()) => {
  const configColumnIndex = convertColumnNameToIndex(CONFIG_COLUMN_NAME);

  const configLabelRowIndex = findConfigLabelRowIndex(ACCOUNTABILITY_PARTNER_COUNT_CONFIG_LABEL, sheet);
  const configLabelCell = sheet.getRange(configLabelRowIndex, configColumnIndex);
  const configCell = getCellRelativeToCell(configLabelCell, 1, 0, sheet);
  const accountabilityPartnerCountConfig = configCell.getValue();

  const currentAccountabilityPartnerCount = findConfigLabelOccurrenceCount(PARTNER_NAME_CONFIG_LABEL, sheet);

  const countDifference = accountabilityPartnerCountConfig - currentAccountabilityPartnerCount;

  if (countDifference < 0 || accountabilityPartnerCountConfig === 0) {
    configCell.setValue(currentAccountabilityPartnerCount);
  } else if (countDifference > 0) {
    const accountabilityPartnersColumnIndex = convertColumnNameToIndex(ACCOUNTABILITY_PARTNER_COLUMN_NAME);

    for (let i = 0; i < countDifference; i++) {
      const newConfigRowsCount = 4;
      const newHabitTableRowsCount = 2;

      const firstAccountabilityPartnerConfigCellRowIndex = configCell.getRowIndex() + 2;
      const newConfigRowsCountOffset = i * newConfigRowsCount;
      const newHabitTableRowsCountOffset = i * newHabitTableRowsCount;
      const minimumRowOffset = firstAccountabilityPartnerConfigCellRowIndex + newConfigRowsCountOffset + newHabitTableRowsCountOffset;

      const emptyCellAfterConfigs = getFirstEmptyCellInColumn(minimumRowOffset, configColumnIndex, sheet);
      sheet.insertRowsAfter(emptyCellAfterConfigs.getRowIndex(), newConfigRowsCount);

      const placeholderName = `${PARTNER_NAME_CONFIG_PLACEHOLDER} ${currentAccountabilityPartnerCount + i + 1}`

      emptyCellAfterConfigs.setValue(`${currentAccountabilityPartnerCount + i + 1}. ${PARTNER_NAME_CONFIG_LABEL}`);
      getCellRelativeToCell(emptyCellAfterConfigs, 1, 0, sheet).setValue(placeholderName).setHorizontalAlignment("right");
      getCellRelativeToCell(emptyCellAfterConfigs, 2, 0, sheet).setValue(HABIT_COUNT_CONFIG_LABEL);
      getCellRelativeToCell(emptyCellAfterConfigs, 3, 0, sheet).setValue(HABIT_COUNT_CONFIG_PLACEHOLDER).setHorizontalAlignment("center");
      sheet.getRange(emptyCellAfterConfigs.getRowIndex(), configColumnIndex, newConfigRowsCount)
        .setBorder(
          true, // Top border
          true, // Left border
          true, // Bottom border
          true, // Right border
          null, // Vertical interior borders
          null, // Horizontal interior borders
          "#000000",
          SpreadsheetApp.BorderStyle.SOLID
        );

      const firstEmptyCellAfterAccountabilityPartners = getFirstEmptyCellInColumn(FIRST_ACCOUNTABILITY_PARTNER_ROW_INDEX, accountabilityPartnersColumnIndex, sheet);
      sheet.insertRowsAfter(firstEmptyCellAfterAccountabilityPartners.getRowIndex(), newHabitTableRowsCount);

      firstEmptyCellAfterAccountabilityPartners.setValue(placeholderName);
      getCellRelativeToCell(firstEmptyCellAfterAccountabilityPartners, 1, 0, sheet).setValue(`${HABIT_PLACEHOLDER} 1`);
      getCellRelativeToCell(firstEmptyCellAfterAccountabilityPartners, 1, 1, sheet).insertCheckboxes();
      sheet.getRange(firstEmptyCellAfterAccountabilityPartners.getRowIndex(), accountabilityPartnersColumnIndex, newHabitTableRowsCount, sheet.getLastColumn())
        .setBorder(
          true, // Top border
          true, // Left border
          true, // Bottom border
          true, // Right border
          null, // Vertical interior borders
          null, // Horizontal interior borders
          "#000000",
          SpreadsheetApp.BorderStyle.SOLID
        );
    }
  }
};

const handleAccountabilityPartnerNameConfig = (newName, sheet = getHabitSheet()) => {
  const columnIndex = convertColumnNameToIndex(ACCOUNTABILITY_PARTNER_COLUMN_NAME);
  const firstEmptyCellAfterAccountabilityPartners = getFirstEmptyCellInColumn(FIRST_ACCOUNTABILITY_PARTNER_ROW_INDEX, columnIndex, sheet);
  const column = sheet.getRange(FIRST_ACCOUNTABILITY_PARTNER_ROW_INDEX, columnIndex, firstEmptyCellAfterAccountabilityPartners.getRowIndex() - FIRST_ACCOUNTABILITY_PARTNER_ROW_INDEX);
  const values = column.getValues().flat();

  const configColumnIndex = convertColumnNameToIndex(CONFIG_COLUMN_NAME);
  const configLabelRowIndex = findConfigLabelRowIndex(PARTNER_NAME_CONFIG_LABEL, sheet);
  const emptyCellAfterConfigs = getFirstEmptyCellInColumn(configLabelRowIndex, configColumnIndex, sheet);
  const configs = sheet.getRange(configLabelRowIndex, configColumnIndex, emptyCellAfterConfigs.getRowIndex());
  const configValues = configs.getValues().flat();

  values.some((value, index) => {
    const currentCell = column.getCell(index + 1, 1);
    const cellOnTheRight = getCellRelativeToCell(currentCell, 0, 1, sheet);
    const dataValidation = cellOnTheRight.getDataValidation();
    const isPartnerNameCell = !dataValidation || dataValidation.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX;
    if (isPartnerNameCell && !configValues.includes(value)) {
      currentCell.setValue(newName);
      return true;
    }

    return false;
  });
};



const handleHabitCountConfig = (editedConfigCell, sheet = getHabitSheet()) => {
  const columnIndex = convertColumnNameToIndex(ACCOUNTABILITY_PARTNER_COLUMN_NAME);
  const firstEmptyCellAfterAccountabilityPartners = getFirstEmptyCellInColumn(FIRST_ACCOUNTABILITY_PARTNER_ROW_INDEX, columnIndex, sheet);
  const column = sheet.getRange(FIRST_ACCOUNTABILITY_PARTNER_ROW_INDEX, columnIndex, firstEmptyCellAfterAccountabilityPartners.getRowIndex() - FIRST_ACCOUNTABILITY_PARTNER_ROW_INDEX);
  const values = column.getValues().flat();

  let habitCount = 0;
  let isPartnerFound = false;
  values.some((_value, index) => {
    const currentCell = column.getCell(index + 1, 1);
    const cellOnTheRight = getCellRelativeToCell(currentCell, 0, 1, sheet);
    const dataValidation = cellOnTheRight.getDataValidation();
    const isPartnerNameCell = !dataValidation || dataValidation.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX;
    if (isPartnerNameCell) {
      const partnerName = getCellRelativeToCell(editedConfigCell, -2, 0, sheet).getValue().toString();
      if(currentCell.getValue().toString() === partnerName) {
        isPartnerFound = true;
      } else if(isPartnerFound) {
        const newHabitCount = editedConfigCell.getValue().toString();
        if (habitCount > newHabitCount || newHabitCount === 0) {
          editedConfigCell.setValue(habitCount);
        } else {
          const countDifference = newHabitCount - habitCount;
          for (let i = 0; i < countDifference; i++) {
            const lastHabitCell = getCellRelativeToCell(currentCell, i - 1, 0, sheet);
            sheet.insertRowAfter(lastHabitCell.getRowIndex());

            const newHabitCell = getCellRelativeToCell(lastHabitCell, 1, 0, sheet);
            newHabitCell.setValue(`${HABIT_PLACEHOLDER} ${habitCount + i + 1}`);
            sheet.getRange(newHabitCell.getRowIndex(), newHabitCell.getColumn() + 2, 1, sheet.getLastColumn()).setValue(null).setDataValidation(null);

            if (i === 0) {
              sheet.getRange(lastHabitCell.getRowIndex(), lastHabitCell.getColumn(), 1, sheet.getLastColumn())
                .setBorder(
                  false, // Top border
                  true, // Left border
                  false, // Bottom border
                  true, // Right border
                  null, // Vertical interior borders
                  null, // Horizontal interior borders
                  "#000000",
                  SpreadsheetApp.BorderStyle.SOLID
                );
            }
            
            if (i === countDifference - 1) {
              sheet.getRange(newHabitCell.getRowIndex(), newHabitCell.getColumn(), 1, sheet.getLastColumn())
                .setBorder(
                  false, // Top border
                  true, // Left border
                  true, // Bottom border
                  true, // Right border
                  null, // Vertical interior borders
                  null, // Horizontal interior borders
                  "#000000",
                  SpreadsheetApp.BorderStyle.SOLID
                );
            }
          }
        }
        return true;
      }
    } else if (isPartnerFound) {
      habitCount++;
    }

    return false;
  });
};

const handleHabitTableChange = (editedRange, sheet = getHabitSheet()) => {
  const value = editedRange.getValue();
  if (value === "" || value === null) {
    const cellOnTheRight = getCellRelativeToCell(editedRange, 0, 1, sheet);
    const dataValidation = cellOnTheRight.getDataValidation();
    const isHabitCell = dataValidation && dataValidation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX;
    if (isHabitCell) {
      sheet.deleteRow(editedRange.getRowIndex());
    }
  } 
};

const handleDateFormatConfig = (sheet = getHabitSheet()) => {
  const configColumnIndex = convertColumnNameToIndex(CONFIG_COLUMN_NAME);
  const configLabelRowIndex = findConfigLabelRowIndex(DATE_FORMAT_CONFIG_LABEL, sheet);

  const configLabelCell = sheet.getRange(configLabelRowIndex, configColumnIndex);
  const configCell = getCellRelativeToCell(configLabelCell, 1, 0, sheet);

  const firstDateColumnIndex = convertColumnNameToIndex(FIRST_DATE_COLUMN_NAME);
  const firstDateCell = sheet.getRange(DATE_ROW_INDEX, firstDateColumnIndex);

  configCell.copyFormatToRange(sheet, firstDateCell.getColumn(), sheet.getLastColumn(), firstDateCell.getRow(), firstDateCell.getLastRow());

  sheet.autoResizeColumns(firstDateCell.getColumn(), sheet.getLastColumn());
  const autoWidth = sheet.getColumnWidth(firstDateColumnIndex);
  const newWidth = autoWidth + 30;
  sheet.setColumnWidths(firstDateColumnIndex, sheet.getLastColumn(), newWidth);
};

const onEdit = (event) => {
  const editedRange = event.range;
  const sheet = editedRange.getSheet();
  const sheetName = PropertiesService.getScriptProperties().getProperty(SHEET_TAB_NAME);

  if (sheet.getName() === sheetName) {
    const configColumnIndex = convertColumnNameToIndex(CONFIG_COLUMN_NAME);

    const isConfigColumnEdited = editedRange.getColumn() === configColumnIndex;
    const isSingleCellEdited = editedRange.getNumRows() === 1 && editedRange.getNumColumns() === 1;

    if (isConfigColumnEdited && isSingleCellEdited && editedRange.getRowIndex() >= FIRST_ACCOUNTABILITY_PARTNER_ROW_INDEX) {
      const cellTextAboveEditedCell = getCellRelativeToCell(editedRange, -1, 0, sheet).getValue().toString();
      const accountabilityPartnersColumnIndex = convertColumnNameToIndex(ACCOUNTABILITY_PARTNER_COLUMN_NAME);
      const firstEmptyCellAfterAccountabilityPartners = getFirstEmptyCellInColumn(FIRST_ACCOUNTABILITY_PARTNER_ROW_INDEX, accountabilityPartnersColumnIndex, sheet);

      switch (true) {
        case cellTextAboveEditedCell.includes(PARTNER_NAME_CONFIG_LABEL):
            handleAccountabilityPartnerNameConfig(editedRange.getValue().toString(), sheet);
          break;

        case cellTextAboveEditedCell === HABIT_COUNT_CONFIG_LABEL:
            handleHabitCountConfig(editedRange, sheet);
          break;

        case editedRange.getRowIndex() <= firstEmptyCellAfterAccountabilityPartners.getRowIndex():
            handleHabitTableChange(editedRange, sheet);
          break;

        default: {
            const accountabilityPartnerCountConfigLabelRowIndex = findConfigLabelRowIndex(ACCOUNTABILITY_PARTNER_COUNT_CONFIG_LABEL, sheet);
            const accountabilityPartnerCountConfigLabelCell = sheet.getRange(accountabilityPartnerCountConfigLabelRowIndex, configColumnIndex);
            const accountabilityPartnerCountConfigCell = getCellRelativeToCell(accountabilityPartnerCountConfigLabelCell, 1, 0, sheet);

            const dateFormatReadyConfigLabelRowIndex = findConfigLabelRowIndex(DATE_FORMAT_READY_CONFIG_LABEL, sheet);
            const dateFormatReadyConfigLabelCell = sheet.getRange(dateFormatReadyConfigLabelRowIndex, configColumnIndex);
            const dateFormatReadyConfigCell = getCellRelativeToCell(dateFormatReadyConfigLabelCell, 1, 0, sheet);

            switch (editedRange.getRow()) {
              case accountabilityPartnerCountConfigCell.getRow():
                  handleAccountabilityPartnerCountConfig(sheet);
                break;

              case dateFormatReadyConfigCell.getRow(): {
                  handleDateFormatConfig(sheet);
                  dateFormatReadyConfigCell.setValue(false);
                break;
              }

              default:
                break;
            }
          break;
        }
      } 
    }
  }
};

/**
 * This function runs on sheet edit, separately from the in-built onEdit function
 * because the onEdit function cannot access the in-built ScriptApp object
 * becasue of authorization reasons but ScriptApp object is needed to update triggers.
 * Its name must be equal to the TIMEZONE_CONFIG_HANDLER_FUNCTION_NAME constant.
 */
const handleTimezoneConfig = (editEvent) => {
  const sheetName = PropertiesService.getScriptProperties().getProperty(SHEET_TAB_NAME);
  const sheet = editEvent.source.getSheetByName(sheetName);

  const { columnStart, columnEnd, rowStart, rowEnd } = editEvent.range;

  const configColumnIndex = convertColumnNameToIndex(CONFIG_COLUMN_NAME);
  const configLabelRowIndex = findConfigLabelRowIndex(TIMEZONE_CONFIG_LABEL, sheet);
  const configLabelCell = sheet.getRange(configLabelRowIndex, configColumnIndex);
  const configCell = getCellRelativeToCell(configLabelCell, 1, 0, sheet);

  const isConfigColumnEdited = columnStart === configColumnIndex;
  const isSingleCellEdited = columnStart === columnEnd && rowStart === rowEnd;
  const isTimezoneConfigEdited = rowStart === configCell.getRow();

  if (isConfigColumnEdited && isSingleCellEdited && isTimezoneConfigEdited) {
    const timezone = configCell.getValue();
    if (UNIQUE_UTC_TIMEZONES.includes(timezone)) {
      updateDailyTrigger(timezone, sheet);
    }
  }
};
