function main(workbook: ExcelScript.Workbook) {

  let cellHasValue: number[][] = [];
  let cellValue: string[] = [];
  let undersizedClassColor: string = "blue";
  let mwfColors: string[] = ["b4f2ac", "dfe8ae", "fce4d6", undersizedClassColor];
  let tthColors: string[] = ["b2bded", "adf0dd", "e9c6f7", undersizedClassColor];
  let assignedRoomColor = "green";
  let assignedAndUnderSizedColor = "red";
  let colorIndex = 0;
  let colorsApplied = false;
  let roomHeadings: string[][] = [];
  let days: string[] = ["M", "T", "W", "Th", "F"]
  var classRoomsTimes: number[][] = [];
  let startingRowIndexForTimes = 2; // it's actually row 3, but they are indexed from 0
  let startingRowIndexForCourseData = 4;
  type colorAndRange = {
    range: ExcelScript.Range;
    color: string;
  }
  let roomColorRanges: colorAndRange[] = [];
  let undersizedClassBuffer = 10;

  type room = {
    name: string;
    capacity: number;
    schedule: Map<string, string[][]>; // e.g. schedule["M"] = monday's schedule (which is times that are filled by class names). The two strings of the tuple are the schedule cells themselves (to be filled with data) and the cell color values.
    colorBlocks: number[][]; // Store the color values for classes that take up blocks in the schedule, for use when we draw it on the Calendar sheet to fill in cell color backgrounds
  }

  let rooms: room[] = []

  var courses = workbook.getWorksheet("Courses");
  var calendar = workbook.getWorksheet("Calendar");
  var roomsSheet = workbook.getWorksheet("Rooms");

  var entireSheet: ExcelScript.Range = calendar.getRange("A1:ZZ1000");

  type UVAClass =
    {
      courseMnemonic: string;
      courseNumber: string;
      courseSection: string;
      day: string;
      enrollment: number;
      startTime: number;
      endTime: number;
      assignedRoom: string;
      uniqueIndex: number;
      rowInDatabase: number;
    };

  let usedRange = courses.getUsedRange();
  let rowCount = usedRange.getRowCount();
  let uvaClasses: UVAClass[] = [];
  let classData = courses.getRange("A4:O" + rowCount as string).getValues();

  // Start your engines!

  getDataFromCoursesSheet();
  createRooms();
  createScheduleBasics();
  fillInManuallyAssignedClasses();
  attemptToMakeSchedule();
  createRoomSchedule();
  applyColorsFromColorBlock();


  function getDataFromCoursesSheet() {
    for (let i = 0; i < classData.length; i++) {
      uvaClasses[i] = {
        courseMnemonic: classData[i][0] as string,
        courseNumber: classData[i][1] as string,
        courseSection: classData[i][2] as string,
        day: classData[i][5] as string,
        enrollment: classData[i][9] as number,
        startTime: classData[i][6] as number,
        endTime: classData[i][7] as number,
        assignedRoom: classData[i][12] as string,
        uniqueIndex: classData[i][14] as number,
        rowInDatabase: i + startingRowIndexForCourseData
      };
    }
  }

  function createRooms() { 

    let roomsRange = roomsSheet.getUsedRange();
    let roomsCount = roomsRange.getRowCount();
    let roomsValues = roomsRange.getValues();
    for (let i = 1; i < roomsCount; i++) { // 1 because first row is heading
      let thisRoom: room = {
        name: roomsValues[i][0] as string,
        capacity: roomsValues[i][1] as number,
        schedule: initializeRoomSchedule(),
        colorBlocks: []
      }

      rooms.push(thisRoom);
    }
    rooms.sort((a, b) => (a.capacity < b.capacity) ? -1 : 1);
  }


  function initializeRoomSchedule(): Map<string, string[][]> { // This fills the arrays with empty values so we can put classes in specific indexes of the schedule
    let schedule: Map<string, string[][]> = new Map<string, string[][]>();
    let defaultValuesM: string[][] = [];
    let defaultValuesTu: string[][] = [];
    let defaultValuesW: string[][] = [];
    let defaultValuesTh: string[][] = [];
    let defaultValuesF: string[][] = [];

    for (let i = 0; i < 53; i++) {
      defaultValuesM[i] = [""];
      defaultValuesTu[i] = [""];
      defaultValuesW[i] = [""];
      defaultValuesTh[i] = [""];
      defaultValuesF[i] = [""];
    }

    schedule.set("M", defaultValuesM);
    schedule.set("T", defaultValuesTu);
    schedule.set("W", defaultValuesW);
    schedule.set("Th", defaultValuesTh);
    schedule.set("F", defaultValuesF);

    return schedule;
  }

  function fillInManuallyAssignedClasses() { // Fill classes that have aready been assigned
    let timeValues = calendar.getRange("A3:A56").getValues();
    for (let i = 0; i < uvaClasses.length; i++) {
      if (uvaClasses[i].assignedRoom != "") {
        let roomIndex = rooms.findIndex(room => room.name == uvaClasses[i].assignedRoom);
        let timeIndex = timeValues.findIndex(time => time[0] == uvaClasses[i].startTime);

        for (let j = timeIndex; j < timeValues.length; j++) {
          if (uvaClasses[i].endTime > timeValues[j][0]) {
            let isRoomAssigned = "yes";
            let isRoomAssignedAndClassUndersized = "both";
            let rowToFill = j;
            if (uvaClasses[i].enrollment + undersizedClassBuffer < rooms[roomIndex].capacity) {
              fillOpenSlot(uvaClasses[i], rowToFill, roomIndex, isRoomAssignedAndClassUndersized);
            }
            else fillOpenSlot(uvaClasses[i], rowToFill, roomIndex, isRoomAssigned);
          }
          else break;
        }
      }
    }
  }

  function attemptToMakeSchedule() {

    let times = calendar.getRange("A3:A56");
    let timeRowCount = times.getRowCount();
    let cellTimeValues = times.getValues() as number[][];

    for (let i = 0; i < uvaClasses.length; i++) {

      if (uvaClasses[i].assignedRoom != "") { // skip manually assigned rooms
        continue;
      }

      for (let j = 0; j < timeRowCount; j++) {
        if (cellTimeValues[j][0] == uvaClasses[i].startTime) {
          let wholeCourseDurationOpenOnCalendar = 1; // boolean as number
          var foundRoomIndex: number;
          let courseDuration = 0;
          let currentRow = j;

          for (let k = 0; k < 10; k++) { // we need to check for the whole duration of the course
            let nextRowChunkToCheck = k;
            if (cellTimeValues[currentRow + nextRowChunkToCheck][0] >= uvaClasses[i].startTime && cellTimeValues[currentRow + nextRowChunkToCheck][0] < uvaClasses[i].endTime) {
              courseDuration++;
            } // If the time row to check is still between the start and end times of the course, we check for it
            else
              break;
          }

          let roomFindAttempt /* Tuple */ = attemptToFindOpenSlot(uvaClasses[i], currentRow, courseDuration);
          let isRoomFound = roomFindAttempt[0] == 1 ? true : false;
          let roomFoundIndex = roomFindAttempt[1];

          if (isRoomFound) { // if we checked the whole duration and it's open
            for (let k = 0; k < courseDuration; k++) {
              let isRoomAlreadyAssigned = "no";
              fillOpenSlot(uvaClasses[i], currentRow + k, roomFoundIndex, isRoomAlreadyAssigned);
            }

          }
        }
      }

    }

  }


  function attemptToFindOpenSlot(uvaClass: UVAClass, row: number, courseDuration: number): [number, number] {

    let foundSpot = 0; // using number as boolean
    var roomIndex: number;
    let daysOfWeek: number[] = []; // we capture which days to populate on the sheet

    for (let i = 0; i < rooms.length; i++) { // Go  through all rooms, which have been ordered by size, and check for open slots
      let truthValues: number[] = []; // We collect a bunch of values and then multiply them together at the end-- if a single zero (0) is in the lot, signifying one of the days is full, it is false
      if (rooms[i].capacity >= uvaClass.enrollment) {

        foundSpot = 1; // we have a room that could hypothetically hold the class, but we need to check each day

        for (let j = 0; j < courseDuration; j++) { // we need to check the whole duration
          if (uvaClass.day.includes("M")) {
            if (rooms[i].schedule.get("M")[row + j][0] == "") {
              truthValues.push(1);
              if (!daysOfWeek.includes(0))
                daysOfWeek.push(0);
            }
            else truthValues.push(0);
          }
          if (uvaClass.day == "TuTh" || uvaClass.day == "T") {

            if (rooms[i].schedule.get("T")[row + j][0] == "") {
              truthValues.push(1);
              if (!daysOfWeek.includes(1))
                daysOfWeek.push(1);
            }
            else truthValues.push(0);
          }
          if (uvaClass.day.includes("W")) {
            if (rooms[i].schedule.get("W")[row + j][0] == "") {
              truthValues.push(1);
              if (!daysOfWeek.includes(2))
                daysOfWeek.push(2);
            }
            else truthValues.push(0);
          }
          if (uvaClass.day.includes("Th")) {
            if (rooms[i].schedule.get("Th")[row + j][0] == "") {
              truthValues.push(1);
              if (!daysOfWeek.includes(3))
                daysOfWeek.push(3);
            }
            else truthValues.push(0);
          }
          if (uvaClass.day.includes("F")) {
            if (rooms[i].schedule.get("F")[row + j][0] == "") {
              truthValues.push(1);
              if (!daysOfWeek.includes(4))
                daysOfWeek.push(4);
            }
            else truthValues.push(0);
          }
        }

      }

      checkAllTruthValues();

      function checkAllTruthValues() {
        for (let j = 0; j < truthValues.length; j++) {
          foundSpot = foundSpot * truthValues[j];
        }
      }


      if (foundSpot == 1) { // If we found a spot, put it in the array for that room's color blocks, and assign the room index for return later
        roomIndex = i;
        for (let j = 0; j < daysOfWeek.length; j++) {
          if (uvaClass.enrollment + undersizedClassBuffer < rooms[roomIndex].capacity) {
            rooms[roomIndex].colorBlocks.push([row, daysOfWeek[j], courseDuration, uvaClass.rowInDatabase, 3]); // the color in index 3 is the color for undersized classes
          }
          else 
            rooms[roomIndex].colorBlocks.push([row, daysOfWeek[j], courseDuration, uvaClass.rowInDatabase, colorIndex]); // We use this data later to add color to the block
        }     
        break;
      }
    }

    return [foundSpot, roomIndex];

  }

  function getNumberOfDay(day: string): number {
    if (day == "M") return 0;
    if (day == "T") return 1;
    if (day == "W") return 2;
    if (day == "Th") return 3;
    if (day == "F") return 4;
  }

  function fillOpenSlot(uvaClass: UVAClass, row: number, foundRoomIndex: number, isRoomAssigned: string) {
    let courseInfo = uvaClass.courseMnemonic + " " + uvaClass.courseNumber + " " + uvaClass.courseSection + " [" + uvaClass.enrollment + "] " + " {" + uvaClass.rowInDatabase + "}";

    function checkIfAlreadyAssigned(dayOfWeek: string) {// If we assigned two or more classes to the same room at the same time
      if (rooms[foundRoomIndex].schedule.get(dayOfWeek)[row][0] != "") {
        throw "You may have assigned this class: " + uvaClass.courseMnemonic + " " + uvaClass.courseNumber + " " + uvaClass.courseSection + " " + uvaClass.rowInDatabase + " to the same room at the same time as this class: " + rooms[foundRoomIndex].schedule.get(dayOfWeek)[row][0];
      }
    }

    function assignScheduleValues(dayOfWeek: string) {
      rooms[foundRoomIndex].schedule.get(dayOfWeek)[row][0] = courseInfo;
      rooms[foundRoomIndex].schedule.get(dayOfWeek)[row].push(isRoomAssigned);
    }

    if (uvaClass.day.includes("M")) {
      checkIfAlreadyAssigned("M");
      assignScheduleValues("M");
    }
    if (uvaClass.day == "TuTh" || uvaClass.day == "T") {
      checkIfAlreadyAssigned("T");
      assignScheduleValues("T");
    }

    if (uvaClass.day.includes("W")) {
      checkIfAlreadyAssigned("W");
      assignScheduleValues("W");
    }
    if (uvaClass.day.includes("Th")) {
      checkIfAlreadyAssigned("Th");
      assignScheduleValues("Th");
    }
    if (uvaClass.day.includes("F")) {
      checkIfAlreadyAssigned("F");
      assignScheduleValues("F");
    }

  }

  function createScheduleBasics() { // Fills out the times 
    entireSheet.clear();
    let hour = 8;
    let mod = 0;
    let minutes = ["00", "15", "30", "45"];


    for (let i = 2; i < 56; i++) {
      let timeCell = entireSheet.getCell(i, 0);
      timeCell.setNumberFormatLocal("hh:mm:ss AM/PM");

      timeCell.setValue(hour.toString() + ":" + minutes[mod].toString());

      mod++;

      if (mod > 3) {
        mod = 0;
        hour++;
      }

    }
  }

  function createRoomSchedule() { // Fills in the classes in rooms

    let spacer = 6; // 5 days + 1 column white space
    console.log("filling schedule");
    for (let i = 0; i < rooms.length; i++) {
      let roomNameCell = calendar.getCell(0, (i * spacer) + 1); // Get every 7th cell, staring in the 2nd column
      roomNameCell.setValue(rooms[i].name);
      roomNameCell.getFormat().getFont().setBold(true);

      function makeColumn(spacerValue: number, dayOfWeek: string, dayChar: string) {
        let courseScheduleData: string[][] = [];
        for (let j = 0; j < rooms[i].schedule.get(dayChar).length; j++) { // separate out the schedule data from any other relevant data
          courseScheduleData.push([""]);
          courseScheduleData[j][0] = rooms[i].schedule.get(dayChar)[j][0];
        }
        let dayCell = entireSheet.getCell(1, (i * spacer) + spacerValue);
        dayCell.setValue(dayOfWeek);
        let targetCell = entireSheet.getCell(2, (i * spacer) + spacerValue);
        let targetRange = targetCell.getResizedRange(rooms[i].schedule.get(dayChar).length - 1, 0);

        targetRange.setValues(courseScheduleData);

        setRoomAssignedColor();
        //setColorsOfRoomsThatAreUndersized();

        function setRoomAssignedColor() {
          for (let k = 0; k < rooms[i].schedule.get(dayChar).length - 1; k++) {
            let isAssigned = rooms[i].schedule.get(dayChar)[k][1] == "yes" ? true : false;
            let isAssignedAndClassUndersized = rooms[i].schedule.get(dayChar)[k][1] == "both" ? true : false;
            if (isAssigned) {
              let coloredCell = entireSheet.getCell(k + startingRowIndexForTimes, (i * spacer) + spacerValue);
              coloredCell.getFormat().getFill().setColor(assignedRoomColor);
            }
            else if (isAssignedAndClassUndersized) {
              let coloredCell = entireSheet.getCell(k + startingRowIndexForTimes, (i * spacer) + spacerValue);
              coloredCell.getFormat().getFill().setColor(assignedAndUnderSizedColor);
            }
          }
        }
          

      }

      setColorsFromColorBlocksBasedOnSpacing();

      makeColumn(1, "Monday", "M");
      makeColumn(2, "Tuesday", "T");
      makeColumn(3, "Wednesday", "W");
      makeColumn(4, "Thursday", "Th");
      makeColumn(5, "Friday", "F");

        
      function setColorsFromColorBlocksBasedOnSpacing() {
        for (let j = 0; j < rooms[i].colorBlocks.length; j++) {
          let blocksToBeColored = rooms[i].colorBlocks;
          blocksToBeColored.sort((a, b) => (a[0] < b[0]) ? -1 : 1); // sort them by row
          let startRow = blocksToBeColored[j][0] + startingRowIndexForTimes;
          let startColumn = blocksToBeColored[j][1] + 1; // One because first column in sheet is times
          let rowsInRange = blocksToBeColored[j][2];
          
          for (let k = j; k >= 0; k--) { // switch colors for new classes
            let currentColumn = blocksToBeColored[j][1];
            let columnToCheckAgainst = blocksToBeColored[k][1]
            let currentClassRowID = blocksToBeColored[j][3];
            let classRowIDToCheckAgainst = blocksToBeColored[k][3];
            let currentBlockColor = blocksToBeColored[j][4];
            let blockColorToCheckAgainst = blocksToBeColored[k][4];
            if (currentColumn == columnToCheckAgainst && 
              currentClassRowID != classRowIDToCheckAgainst) { // if it's in the same column, but a different class ID
              
              switch(blockColorToCheckAgainst) {
                case 2: 
                  currentBlockColor = 0;
                  break;
                case 1: 
                  currentBlockColor = 2;
                  break;
                case 0: 
                  currentBlockColor = 1;
                  break;
                case 3:
                  currentBlockColor = 3;
                  break;
              }
             
              blocksToBeColored[j][4] = currentBlockColor;
              break;
            }
          }

          let colorIndex = blocksToBeColored[j][4];
          var classColor: string;

          let dayOfWeekBasedOnColumnNumber = blocksToBeColored[j][1];
          if (dayOfWeekBasedOnColumnNumber == 0 || dayOfWeekBasedOnColumnNumber == 2 || dayOfWeekBasedOnColumnNumber == 4)
            classColor = mwfColors[colorIndex];// different colors for MWF TuTh
          else classColor = tthColors[colorIndex]

          let colorRange = calendar.getRangeByIndexes(startRow, startColumn + (i * spacer), rowsInRange, 1);

          roomColorRanges.push({ range: colorRange, color: classColor });

          for (let k = j; k >= 0; k--) { // if we switched colors on the last day of a class (because it had a class above it), go back and make every day of that class that color
            let currentClassRowID = blocksToBeColored[j][3];
            let classRowIDToCheckAgainst = blocksToBeColored[k][3];
            if (currentClassRowID == classRowIDToCheckAgainst) { // check for same class
              blocksToBeColored[k][4] = blocksToBeColored[j][4]; // make sure they are all the same color in the array
              // Set the range based on the previous class locations in the array
              let previousClassColorBlockRow = blocksToBeColored[k][0] + startingRowIndexForTimes;
              let previousClassColorBlockColumn = blocksToBeColored[k][1]; 
              let previousClassColorBlockColumnInSheet = blocksToBeColored[k][1] + (i * spacer) + 1; 
              let previousClassColorBlockDuration = blocksToBeColored[k][2];
              let amendedColorRange = calendar.getRangeByIndexes(previousClassColorBlockRow, previousClassColorBlockColumnInSheet, previousClassColorBlockDuration, 1);
              // And set the colors based on day
              if (previousClassColorBlockColumn == 0 || previousClassColorBlockColumn == 2 || previousClassColorBlockColumn == 4)
                classColor = mwfColors[colorIndex];// different colors for MWF TuTh
              else classColor = tthColors[colorIndex]

              roomColorRanges.push({ range: amendedColorRange, color: classColor });
            }
          }
        }
      }

      


    }

  }

  function applyColorsFromColorBlock() {
    console.log("adding colors"); // This clears the memory buffer

    for (let i = 0; i < roomColorRanges.length; i++) {
      let color = roomColorRanges[i].color;
      roomColorRanges[i].range.getFormat().getFill().setColor(color);
    }
  }
  

}



