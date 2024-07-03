function main(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getWorksheet("ITSM Manage Firmware Template");

    const snColumn = 3; // Column D
    const ipColumn = 5; // Column F

    let usedRange = sheet.getUsedRange();
    let rowCount = usedRange.getRowCount();
    let errorMessages: string[] = [];

    for (let i = 1; i < rowCount; i++) {
        if (!usedRange.getCell(i, snColumn).getValue()) {
            errorMessages.push(`Row ${i + 1}: Column D cannot be empty \n`);
        }
        if (!usedRange.getCell(i, ipColumn).getValue()) {
          errorMessages.push(`Row ${i + 1}: Column F cannot be empty \n`);
        }
    }

    if (errorMessages.length > 0) {
      try{
        sheet = workbook.addWorksheet("ValidationErrors");
      } catch (exception){
        sheet = workbook.getWorksheet("ValidationErrors");
      }
      
      sheet.activate();
      let selectedCell = workbook.getActiveCell();
      console.log(errorMessages);
      console.log("Validation failed. Errors stored in named range 'ValidationErrors'.");
      for (let i = 0; i < errorMessages.length; i++){
        sheet.getCell(i, 0).setValue(errorMessages[i]);
        //sheet.getCell(i,0).getFormat()
      }
      
    } else {
        console.log("All validations passed");
    }
}
