function main(workbook: ExcelScript.Workbook) {
    // This locks a class on the calendar
    let cell = workbook.getActiveCell();
    let worksheet = workbook.getActiveWorksheet();
    if (worksheet.getName() != "Calendar") {
      throw "You need to do this on the Calendar worksheet";
    }
    let courses = workbook.getWorksheet("Courses");
    let setRoomName: string = "";
    var roomStartColumn: number;
    let roomColor: string = "green";
    var rowNumber: number;

    getRoomNameForCurrentlySelectedClassCell();
    setRoomInCoursesSheet();
    colorTheClassOnCalendarSheet();

    function getRoomNameForCurrentlySelectedClassCell() {
      for (let i = 0; i < 6; i++) {
        let columnIndex = cell.getColumnIndex();
        let roomName = worksheet.getCell(0, columnIndex - i).getValue();
        if (roomName != "") {
          setRoomName = roomName as string;
          roomStartColumn = columnIndex - i;
          break;
        }
      }
    }

    function setRoomInCoursesSheet() {
      let cellValue = cell.getValue().toString();
      rowNumber = +cellValue.slice(cellValue.indexOf("{") + 1, cellValue.indexOf("}")); // The row is listed on the Calendar sheet between {}
      let roomRange = courses.getRange("M1:M1000").getUsedRange();
      
      roomRange.getCell(rowNumber - 1, 0).setValue(setRoomName);
    }
    
    function colorTheClassOnCalendarSheet() {
      let rangeToCheck = worksheet.getRangeByIndexes(0, roomStartColumn, 53, 5).getUsedRange();
      let rowCount = rangeToCheck.getRowCount();
      let columnCount = rangeToCheck.getColumnCount();
      let values = rangeToCheck.getValues();
      for (let i = 0; i < rowCount; i++) {
        for (let j = 0; j < columnCount; j++) {
          let value = values[i][j] as string;
          if (value.includes("{" + rowNumber.toString() + "}")) {
            rangeToCheck.getCell(i,j).getFormat().getFill().setColor(roomColor);
          }
        }
      } 
    }
}