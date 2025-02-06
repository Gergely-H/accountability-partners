const updateDailyTrigger = (timezone = DEFAULT_TIMEZONE) => {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === NEW_DAY_HANDLER_FUNCTION_NAME) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger(NEW_DAY_HANDLER_FUNCTION_NAME)
    .timeBased()
    .atHour(0)
    .inTimezone(timezone)
    .everyDays(1)
    .create();
};

const createOnEditTriggerForTimezoneConfig = () => {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === TIMEZONE_CONFIG_HANDLER_FUNCTION_NAME && triggers[i].getTriggerType() === ScriptApp.EventType.ON_EDIT) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  const spreadsheet = getSpreadsheet();

  ScriptApp.newTrigger(TIMEZONE_CONFIG_HANDLER_FUNCTION_NAME)
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();
};
