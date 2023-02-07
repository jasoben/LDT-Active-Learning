let cellHasValue: number[][] = [];
let cellValue: string[] = [];
let colors: string[] = ["b4f2ac", "dfe8ae", "b2bded", "adf0dd"]
let colorIndex = 0;

function main(workbook: ExcelScript.Workbook) {
  // Your code here

  var courses = workbook.getWorksheet("Courses");

  var calendar = workbook.getActiveWorksheet();

  if (calendar.getName() == "Courses")
  {
    throw "don't run this on the courses page"
  }

  var calendarRange = calendar.getRange("B2:M200");
  calendarRange.clear();

  var day: string = calendar.getRange("N2").getValue() as string;
  var enrollMin: number = calendar.getRange("O2").getValue() as number;
  var enrollMax: number = calendar.getRange("P2").getValue() as number;

  switch (day) {
    case "M":
      calendar.getRange("A1").setValue("Monday");
      break;
    case "T":
      calendar.getRange("A1").setValue("Tuesday");
      break;
    case "W":
      calendar.getRange("A1").setValue("Wednesday");
      break;
    case "Th":
      calendar.getRange("A1").setValue("Thursday");
      break;
    case "F":
      calendar.getRange("A1").setValue("Friday");
      break;
  }

  // Get data
  type UVAClass = 
  {
    courseMnemonic: string;
    courseNumber: string;
    day: string;
    enrollment: number;
    startTime: number;
    endTime: number;
    assignedRoom: string;
    uniqueIndex: number;
  };

  let usedRange = courses.getUsedRange();
  let rowCount = usedRange.getRowCount();
  let uvaClasses : UVAClass[] = [];
  let classData = courses.getRange("A4:O" + rowCount as string).getValues();
  for (let i = 0; i < classData.length; i++)
  {
    uvaClasses[i] = {courseMnemonic: classData[i][0] as string, 
      courseNumber: classData[i][1] as string,
      day: classData[i][5] as string,
      enrollment: classData[i][9] as number,
      startTime: classData[i][6] as number,
      endTime: classData[i][7] as number,
      assignedRoom: classData[i][12] as string,
      uniqueIndex: classData[i][14] as number};
  }

  let calendarTimes: number[][] = calendar.getRange("A2:A54").getValues() as number[][];

  let roomHeadings: string[][] = calendar.getRange("B1:J1").getValues() as string[][];

  console.log(roomHeadings[0].length);
  // Look at days to see if they contain the day of the week
  // we define in N2 on the Calendar sheet, and the enrollment
  // is within the range we set in O2 and P2

  // Check room assignments in data, and if they aren't 
  // on the calendar, make them

  for (let i = 0; i < uvaClasses.length; i++) 
  {
    if (uvaClasses[i].day.includes(day) && CheckTuesdays(uvaClasses[i].day, i) && uvaClasses[i].enrollment <= enrollMax
      && uvaClasses[i].enrollment >= enrollMin &&
      uvaClasses[i].assignedRoom != "") 
      {

        if (!roomHeadings[0].includes(uvaClasses[i].assignedRoom)) {
          for (let j = 0; j < roomHeadings[0].length; j++) {
            if (roomHeadings[0][j] == "") {
              roomHeadings[0][j] = uvaClasses[i].assignedRoom;
              calendar.getCell(0, j + 1).setValue(uvaClasses[i].assignedRoom);
              break;
            }
          }
        }
      }
    }

    function CheckTuesdays(classDay : string, i : number) : boolean
    { 
      if (day == "T" && classDay == "Th")
      {
        return false;
      }
      else return true;
    }

  // Go through first to check for assigned rooms
  for (let u = 0; u < roomHeadings[0].length; u++) {
    for (let i = 0; i < uvaClasses.length; i++) {
      if (uvaClasses[i].day.includes(day) && CheckTuesdays(uvaClasses[i].day, i) && uvaClasses[i].enrollment <= enrollMax
        && uvaClasses[i].enrollment >= enrollMin &&
        uvaClasses[i].assignedRoom == roomHeadings[0][u] &&
        uvaClasses[i].assignedRoom != "") {

        CheckCalendar(i, u + 1, true);

      }
    }
  }

  // Then again for un-assigned rooms
  for (let i = 0; i < uvaClasses.length; i++) {
    if (uvaClasses[i].day.includes(day) && CheckTuesdays(uvaClasses[i].day, i) && uvaClasses[i].enrollment <= enrollMax
      && uvaClasses[i].enrollment >= enrollMin &&
      uvaClasses[i].assignedRoom == "") {
      
      CheckCalendar(i, 1, false);

    }
  }

  function CheckCalendar(i: number, columnNumber: number, checkAssignedOverlap: boolean) {
    let calendarColumnNumber = columnNumber;
    CheckCalendarTimes(uvaClasses[i].startTime, uvaClasses[i].endTime, calendarTimes, calendar, uvaClasses[i].courseMnemonic, uvaClasses[i].courseNumber, uvaClasses[i].enrollment, calendarColumnNumber, checkAssignedOverlap, uvaClasses[i].uniqueIndex);
    colorIndex++; // We use this to change the fill colors on different classes
    if (colorIndex > colors.length - 1) {
      colorIndex = 0;
    }
  }

}
function CheckCalendarTimes(startTime: number, endTime: number, calendarTimes: number[][], calendar: ExcelScript.Worksheet, courseMnemonic: string, courseNumber: string, enrollment: number, calendarColumnNumber: number, checkAssignedOverlap: boolean, uniqueIndex: number) {
  let adjusterNumber = 0;
  let currentCourseTimeBlock = 0; //Used to determine where one 
  // course ends and the other starts

  // Sort the already assigned blocks so we can iterate through them
  // easily in the for loop below

  var sortedCellHasValue = cellHasValue.sort((a, b) => {
    if (a[1] === b[1]) {
      return 0;
    }
    else {
      return (a[1] < b[1]) ? -1 : 1;
    }
  });
  var overlapFound: boolean = false;

  // First go through to see if any times overlap with existing data
  
  for (let i = 0; i < calendarTimes.length; i++) {
    
    let tempAdjusterNumber = 0;
    
    if (startTime <= calendarTimes[i][0] && endTime >= calendarTimes[i][0]) {
      // We use the sorted value here because these used 
      // cells build on top of each other

      for (let j = 0; j < sortedCellHasValue.length; j++) {
        if (sortedCellHasValue[j][0] == i + 1 && sortedCellHasValue[j][1] == calendarColumnNumber + tempAdjusterNumber) {
            if (checkAssignedOverlap) {
              overlapFound = true;
            }
            else {
              tempAdjusterNumber++;
            }
          }
        }
      }

      if (tempAdjusterNumber > adjusterNumber)
      {
        adjusterNumber = tempAdjusterNumber;
      }

    }

    
    for (let i = 0; i < calendarTimes.length; i++) {
      if (startTime <= calendarTimes[i][0] && endTime >= calendarTimes[i][0]) {

        let cell = calendar.getCell(i + 1, calendarColumnNumber + adjusterNumber);
        
        // For assigned rooms where the courses overlap
        if (overlapFound) {
          let newCellValue = LookupStoredCellValue(i + 1, calendarColumnNumber + adjusterNumber, courseMnemonic + " " + courseNumber + ": " + enrollment + " [" + uniqueIndex + "]");
          cell.setValue(newCellValue);
          cell.getFormat().getFill().setColor("red");
        }
        // For unassigned rooms
        else {
          cell.setValue(courseMnemonic + " " + courseNumber + ": " + enrollment + " [" + uniqueIndex + "]");
          cell.getFormat().getFill().setColor(colors[colorIndex]);
        }


        cell.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.dashDot);
        if (currentCourseTimeBlock < 1) {
          cell.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.dashDot);
        }
        currentCourseTimeBlock++;

        cellHasValue.push([i + 1, calendarColumnNumber + adjusterNumber]);
        cellValue.push(courseMnemonic + " " + courseNumber + ": " + enrollment + " [" + uniqueIndex + "]");
      }
    }
    function LookupStoredCellValue(x: number, y:number, currentString: string)
    {
      for(let i = 0; i < sortedCellHasValue.length; i++)
      {
        if (sortedCellHasValue[i][0] == x && sortedCellHasValue[i][1] == y)
        {
          cellValue[i] = cellValue[i] + " / " + currentString; 
          return cellValue[i];
        }
      }
    }
  }
