$(document).ready(function() {var formatter = new CucumberHTML.DOMFormatter($('.cucumber-report'));formatter.uri("C:/Users/VENKATRAMAN/workspace/ExcelManipulation/ExcelManipulation/src/main/java/feature/excel_manipulation.feature");
formatter.feature({
  "line": 1,
  "name": "Excel Manipulation",
  "description": "",
  "id": "excel-manipulation",
  "keyword": "Feature"
});
formatter.scenario({
  "line": 2,
  "name": "Excel Manipulation Scenario",
  "description": "",
  "id": "excel-manipulation;excel-manipulation-scenario",
  "type": "scenario",
  "keyword": "Scenario"
});
formatter.step({
  "line": 3,
  "name": "Excel file value copy and fetch in new Excel File",
  "keyword": "Given "
});
formatter.step({
  "line": 4,
  "name": "Convert MNL to AEST",
  "keyword": "Then "
});
formatter.step({
  "line": 5,
  "name": "Create new Column and find time taken",
  "keyword": "And "
});
formatter.step({
  "line": 6,
  "name": "Remove Dupliacte VCI-Codes and their respective Rows",
  "keyword": "Then "
});
formatter.match({
  "location": "StepDefinition.excel_file_value_copy_and_fetch_in_new_Excel_File()"
});
formatter.result({
  "duration": 1770117543,
  "status": "passed"
});
formatter.match({
  "location": "StepDefinition.convert_MNL_to_AEST()"
});
formatter.result({
  "duration": 51416143,
  "status": "passed"
});
formatter.match({
  "location": "StepDefinition.create_new_Column_and_find_time_taken()"
});
formatter.result({
  "duration": 44975930,
  "status": "passed"
});
formatter.match({
  "location": "StepDefinition.remove_Dupliacte_VCI_Codes_and_their_respective_Rows()"
});
formatter.result({
  "duration": 757232228,
  "status": "passed"
});
});