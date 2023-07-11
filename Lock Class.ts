function main(workbook: ExcelScript.Workbook) {
    // Your code here
    let cell = workbook.getActiveCell();
    let worksheet = workbook.getActiveWorksheet();
    if (worksheet.getName() != "Calendar") {
      throw "You need to do this on the Calendar worksheet";
    }
    let courses = workbook.getWorksheet("Courses");
    let setRoomName: string = "";
    var roomStartColumn: number;

    for (let i = 0; i < 6; i++) {
      let columnIndex = cell.getColumnIndex();
      let roomName = worksheet.getCell(0, columnIndex - i).getValue();
      if (roomName != "") {
        setRoomName = roomName as string;
        roomStartColumn = columnIndex - i;
        break;
      }
    }


    let cellValue = cell.getValue().toString();
    let rowNumber: number = +cellValue.slice(cellValue.indexOf("{") + 1, cellValue.indexOf("}"));
    let roomRange = courses.getRange("M1:M1000").getUsedRange();
    
    roomRange.getCell(rowNumber - 1, 0).setValue(setRoomName);
    
    let rangeToCheck = worksheet.getRangeByIndexes(0, roomStartColumn, 53, 5).getUsedRange();
    let rowCount = rangeToCheck.getRowCount();
    let columnCount = rangeToCheck.getColumnCount();
    let values = rangeToCheck.getValues();
    for (let i = 0; i < rowCount; i++) {
      for (let j = 0; j < columnCount; j++) {
        let value = values[i][j] as string;
        if (value.includes("{" + rowNumber.toString() + "}")) {
          rangeToCheck.getCell(i,j).getFormat().getFill().setColor("red");
        }
      }
      
    } 
}