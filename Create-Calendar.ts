function main(workbook: ExcelScript.Workbook) {

  // Values for times and days
  let startingRowIndexForTimes = 2; // Which row is 800AM on in the chart? (It's actually row 3, but they are indexed from 0)

  // Rooms
  type room = {
    name: string;
    capacity: number;
    schedule: Map<string, classAndColor[]>; // e.g. schedule["M"] = monday's schedule (which is times that are filled by class names). The two strings of the tuple are the schedule cells themselves (to be filled with data) and the cell color values.
    colorBlocks: number[][]; // Store the color values for classes that take up blocks in the schedule, for use when we draw it on the Calendar sheet to fill in cell color backgrounds
  }
  let rooms: room[] = []

  // Worksheets
  var courses = workbook.getWorksheet("Courses");
  var calendar = workbook.getWorksheet("Calendar");
  var roomsSheet = workbook.getWorksheet("Rooms");

  var entireSheet: ExcelScript.Range = calendar.getRange("A1:ZZ1000");

  // Courses
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
      uniqueIndex: number; // The unique index number assigned in the "courses" sheet
      rowInDatabase: number; // The row in the "courses" sheet where this course lives; note, this and the index aren't always the same, depending on how that sheet is sorted
    };

  let usedRange = courses.getUsedRange();
  let rowCount = usedRange.getRowCount();
  let uvaClasses: UVAClass[] = [];

  //Find out where the data starts on the "courses" sheet; since the first column is "Mnemonic", we look for that word 
  //in the cell and know that is the header row, so the data starts one row below that
  let startingRowIndexForCourseData: number = 0;
  let textValues = courses.getRange("A1:A10").getValues()
  for (let i = 0; i < textValues.length; i++) {
    if (textValues[i][0] == "Mnemonic") {
      startingRowIndexForCourseData = i + 1;
    }
  }
  let startingRowNumberForCourseData = startingRowIndexForCourseData + 1; // The row number is plus one since the index is zero indexed
  let classData = courses.getRange("A" + startingRowNumberForCourseData as string + ":O" + rowCount as string).getValues();

  // Undersized classes
  let undersizedClassColor: string = "yellow"; // When courses get placed in rooms that are too large (based on the undersizedClassBuffer, below) we mark them this color
  let undersizedClassBuffer = .67; // Classes with enrollment that are less than this fraction of the room size are undersized

  // Coloring the chart
  let mwfColors: string[] = ["b4f2ac", "dfe8ae", "fce4d6", undersizedClassColor]; // The color blocks we use for MWF classes-- for ease of reading the chart
  let tthColors: string[] = ["b2bded", "adf0dd", "e9c6f7", undersizedClassColor]; // same as above, but for TTh
  let colorIndex = 0; // A number we use to track which set of colors to use: mwf or tth
  let assignedRoomColor = "dbdbdb"; // Color for when we have "locked" or assigned a room (possibly redundant)
  let assignedAndUnderSizedColor = "red"; // Color for when it is "locked" or assigned, and undersized (possibly redundant)
  let manuallyAssignedRoomColor = "green" // Color for when we have "locked" or assigned a room (possibly redundant)
  let manuallyAssignedUndersizedRoomColor = "6161FE" // Color for when it is "locked" or assigned, and undersized (possibly redundant)
  let conflictColor = "red" // Color for when it is "locked" or assigned, and undersized (possibly redundant)
  type colorAndRange = {
    range: ExcelScript.Range;
    color: string;
  } // This is how we store "color blocks" that we then fill in later. 
  let roomColorRanges: colorAndRange[] = []; // The array that stores the colorAndRange types

  type classAndColor = {
    classData: string;
    color: string;
  }


  // Start your engines!

  getDataFromCoursesSheet(); // Grab the data 
  createRooms(); // Create the rooms data
  createScheduleBasics(); // Create the schedule to be filled
  fillInManuallyAssignedClasses(); // Fill in any classes we have already "locked" or assigned
  attemptToMakeSchedule(); // Attempt to fill in courses into the schedule based on size, availability, etc.
  createRoomSchedule(); // Create the schedule for each room based on the above schedule
  applyColorsFromColorBlock(); // Assign the colors to the resulting chart on the "calendar" sheet


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
    rooms.sort((a, b) => (a.capacity < b.capacity) ? -1 : 1); // We sort by capacity because we fill from smallest to largest
  }


  function initializeRoomSchedule(): Map<string, classAndColor[]> { // This fills the arrays with empty values so we can put classes in specific indexes of the schedule
    let schedule: Map<string, classAndColor[]> = new Map<string, classAndColor[]>(); // Where the first string is day of the week, and then the 
      //class data per cell and color data per cell
    let defaultValuesM: classAndColor[] = [];
    let defaultValuesTu: classAndColor[] = [];
    let defaultValuesW: classAndColor[] = [];
    let defaultValuesR: classAndColor[] = [];
    let defaultValuesF: classAndColor[] = [];

    for (let i = 0; i < 53; i++) {
      defaultValuesM[i].classData = "";
      defaultValuesTu[i].classData = "";
      defaultValuesW[i].classData = "";
      defaultValuesR[i].classData = "";
      defaultValuesF[i].classData = "";

      defaultValuesM[i].color = "";
      defaultValuesTu[i].color = "";
      defaultValuesW[i].color = "";
      defaultValuesR[i].color = "";
      defaultValuesF[i].color = "";
    }

    schedule.set("M", defaultValuesM);
    schedule.set("T", defaultValuesTu);
    schedule.set("W", defaultValuesW);
    schedule.set("R", defaultValuesR);
    schedule.set("F", defaultValuesF);

    return schedule;
  }

  function fillInManuallyAssignedClasses() { // Fill classes that have aready been assigned
    let timeValues = calendar.getRange("A3:A56").getValues();
    for (let i = 0; i < uvaClasses.length; i++) {
      // If the room has been assigned, we find that room in our rooms array and the start time in our timeValues array, 
      // then fill the slot in our rooms.schedule array, noting the special colors for "locked" or assigned rooms
      if (uvaClasses[i].assignedRoom != "") {
        let roomIndex = rooms.findIndex(room => room.name == uvaClasses[i].assignedRoom);

        let timeIndex = timeValues.findIndex(time => time[0] == uvaClasses[i].startTime);

        for (let j = timeIndex; j < timeValues.length; j++) {
          if (uvaClasses[i].endTime > timeValues[j][0]) {
            // These values "manual-yes" and "manual-both" are read later when we are filling in colors
            let isRoomManuallyAssigned = "manual-yes";
            let isRoomManuallyAssignedAndClassUndersized = "manual-both";
            let rowToFill = j;
            // Check if it is under capacity
            if (uvaClasses[i].enrollment < rooms[roomIndex].capacity * undersizedClassBuffer) {
              fillOpenSlot(uvaClasses[i], rowToFill, roomIndex, isRoomManuallyAssignedAndClassUndersized);
            }
            else fillOpenSlot(uvaClasses[i], rowToFill, roomIndex, isRoomManuallyAssigned);
          }
          else break;
        }
      }
    }
  }

  function attemptToMakeSchedule() {

    // The times we are using for classes
    let times = calendar.getRange("A3:A56");
    let timeRowCount = times.getRowCount();
    let cellTimeValues = times.getValues() as number[][];

    for (let i = 0; i < uvaClasses.length; i++) {

      if (uvaClasses[i].assignedRoom != "") { // skip manually assigned rooms
        continue;
      }

      for (let j = 0; j < timeRowCount; j++) {
        if (cellTimeValues[j][0] == uvaClasses[i].startTime) {
          let wholeCourseDurationOpenOnCalendar = 1; // boolean as number; we check every day and time and then multiply all these values together-- if there is a zero for "F" it won't place the course in MW
          var foundRoomIndex: number; // The index value of the open room in ours rooms array
          let courseDuration = 0; // How many rows is the duration of the course?
          let currentRow = j; // The current row we are checking 

          // This loop looks at the times and decides how many rows we need to check; we delineate by :15 time chunks, but if we did :30 or :10 this would be
          // a different number of rows
          for (let k = 0; k < 10; k++) { // we need to check for the whole duration of the course
            let nextRowChunkToCheck = k;
            if (cellTimeValues[currentRow + nextRowChunkToCheck][0] >= uvaClasses[i].startTime && cellTimeValues[currentRow + nextRowChunkToCheck][0] < uvaClasses[i].endTime) {
              courseDuration++;
            } // If the time row to check is still between the start and end times of the course, we check for it 
            else
              break;
          }

          let roomFindAttempt /* Tuple */ = attemptToFindOpenSlot(uvaClasses[i], currentRow, courseDuration);  // We send the row and duration data to this function
          // which tries to then check against the rooms.schedule arrays in our rooms array; we don't want to send the raw time values because these
          // are meaningless to the array, which only cares about index values
          let isRoomFound = roomFindAttempt[0] == 1 ? true : false; // Since the function returns a tuple, we look at the first value of the tuple; 
          // if we succeeded, then we found a room
          let roomFoundIndex = roomFindAttempt[1]; // The second value of the tuple is the index of the room in the rooms array

          if (isRoomFound) { // if we checked the whole duration and it's open
            for (let k = 0; k < courseDuration; k++) {
              let isRoomAlreadyAssigned = "no"; // This value is for manual assignment; it may seem counterintuitive that we are "filling" a slot and saying 
              // it hasn't been assigned, but this is to track whether it was manually assigned (i.e. "locked") or not
              fillOpenSlot(uvaClasses[i], currentRow + k, roomFoundIndex, isRoomAlreadyAssigned);
            }

          }
        }
      }

    }

  }


  // This attempts to find an open slot in all our rooms.schedule arrays within the rooms array
  function attemptToFindOpenSlot(uvaClass: UVAClass, row: number, courseDuration: number): [number, number] {

    let foundSpot = 0; // using number as boolean since we check all days; if MW are 1 and F is zero, then MWF is zero since they are multiplied at the end; 
    // this way we can check all days to verify the whole schedule is open
    var roomIndex: number;
    let daysOfWeek: number[] = []; // we capture which days to populate on the sheet; this array let's us know which days we're dealing with for this course

    for (let i = 0; i < rooms.length; i++) { // Go  through all rooms, which have been ordered by size, and check for open slots
      let truthValues: number[] = []; // We collect a bunch of values and then multiply them together at the end-- if a single zero (0) is in the lot, signifying one of the days is full, it is false
      if (rooms[i].capacity >= uvaClass.enrollment) {

        foundSpot = 1; // we have a room that could hypothetically hold the class, but we need to check each day

        // The rooms array contains rooms.schedule arrays that then contain dictionaries with key values for each day "M", "T", "W", etc.
        // Within those are arrays for the slots; importantly, the rooms.schedules don't know or care about times, they only care about slots; 
        // times are handled in attemptToMakeSchedule function and converted to index values to check slots in these arrays
        for (let j = 0; j < courseDuration; j++) { // we need to check the whole duration; again, duration is slots, not actual times, which this array doesn't 
          // know nor care about
          if (uvaClass.day.includes("M")) {
            if (rooms[i].schedule.get("M")[row + j][0] == "") {
              truthValues.push(1);
              if (!daysOfWeek.includes(0)) // If the array doesn't already contain 0 (signifying Monday as the 0 column in the chart), put in 0
                daysOfWeek.push(0);
            }
            else truthValues.push(0);
          }
          if (uvaClass.day.includes("T")) {

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
          if (uvaClass.day.includes("R")) {
            if (rooms[i].schedule.get("R")[row + j][0] == "") {
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

      // Multiply the values all together to make sure we don't have a conflict on one day of the week, e.g. MW are free, but F is not
      function checkAllTruthValues() {
        for (let j = 0; j < truthValues.length; j++) {
          foundSpot = foundSpot * truthValues[j];
        }
      }


      if (foundSpot == 1) { // If we found a spot, put it in the array for that room's color blocks, and assign the room index for return later
        roomIndex = i;
        for (let j = 0; j < daysOfWeek.length; j++) {
          if (uvaClass.enrollment < rooms[roomIndex].capacity * undersizedClassBuffer) {
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


  function fillOpenSlot(uvaClass: UVAClass, row: number, foundRoomIndex: number, isRoomAssigned: string) {
    let courseInfo = uvaClass.courseMnemonic + " " + uvaClass.courseNumber + " " + uvaClass.courseSection + " [" + uvaClass.enrollment + "] " + " <" + uvaClass.rowInDatabase + ">" + " {" + uvaClass.uniqueIndex + "}";

    function checkIfAlreadyAssigned(dayOfWeek: string) {// If we assigned two or more classes to the same room at the same time
      if (rooms[foundRoomIndex].schedule.get(dayOfWeek)[row][0] != "") {
        let currentInfoInRoom = rooms[foundRoomIndex].schedule.get(dayOfWeek)[row][0];
        courseInfo = currentInfoInRoom + " /-/ " + courseInfo;
        console.log("assigned");
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
    if (uvaClass.day.includes("T")) {
      checkIfAlreadyAssigned("T");
      assignScheduleValues("T");
    }

    if (uvaClass.day.includes("W")) {
      checkIfAlreadyAssigned("W");
      assignScheduleValues("W");
    }
    if (uvaClass.day.includes("R")) {
      checkIfAlreadyAssigned("R");
      assignScheduleValues("R");
    }
    if (uvaClass.day.includes("F")) {
      checkIfAlreadyAssigned("F");
      assignScheduleValues("F");
    }

  }

  function createScheduleBasics() { // Fills out the times on the chart
    entireSheet.clear();
    let hour = 8;
    let mod = 0;
    let minutes = ["00", "15", "30", "45"];

    for (let i = 2; i < 56; i++) {
      let timeCell = entireSheet.getCell(i, 0);
      timeCell.setNumberFormatLocal("hh:mm:ss AM/PM");

      timeCell.setValue(hour.toString() + ":" + minutes[mod].toString());

      // modulus for cycling through time chunks see "minutes" array above
      mod++;

      if (mod > 3) {
        mod = 0;
        hour++;
      }

    }
  }

  function createRoomSchedule() { // Fills in the classes in rooms; creates the actual chart

    let spacer = 6; // 5 days + 1 column white space
    console.log("filling schedule");
    for (let i = 0; i < rooms.length; i++) {
      let roomNameCell = calendar.getCell(0, (i * spacer) + 1); // Get every 7th cell, staring in the 2nd column
      roomNameCell.setValue(rooms[i].name);
      roomNameCell.getFormat().getFont().setBold(true);

      makeColumn(1, "Monday", "M");
      makeColumn(2, "Tuesday", "T");
      makeColumn(3, "Wednesday", "W");
      makeColumn(4, "Thursday", "R");
      makeColumn(5, "Friday", "F");

      // Make the columns in the room chart on the Calendars sheet
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

        setColorsFromColorBlocksBasedOnSpacing();
        setRoomAssignedColor();

        // Set colors for rooms that have been manually assigned
        function setRoomAssignedColor() {
          for (let k = 0; k < rooms[i].schedule.get(dayChar).length - 1; k++) {
            let isAssigned = rooms[i].schedule.get(dayChar)[k][1] == "yes" ? true : false;
            let isAssignedAndClassUndersized = rooms[i].schedule.get(dayChar)[k][1] == "both" ? true : false;
            let isManuallyAssigned = rooms[i].schedule.get(dayChar)[k][1] == "manual-yes" ? true : false;
            let isManuallyAssignedAndUndersized = rooms[i].schedule.get(dayChar)[k][1] == "manual-both" ? true : false;
            let hasConflict = rooms[i].schedule.get(dayChar)[k][0].search("/-/") > 0 ? true : false; 
            if (isAssigned) {
              let coloredCell = entireSheet.getCell(k + startingRowIndexForTimes, (i * spacer) + spacerValue);
              coloredCell.getFormat().getFill().setColor(assignedRoomColor);
            }
            else if (isAssignedAndClassUndersized) {
              let coloredCell = entireSheet.getCell(k + startingRowIndexForTimes, (i * spacer) + spacerValue);
              coloredCell.getFormat().getFill().setColor(assignedAndUnderSizedColor);
            }
            else if (isManuallyAssigned) {
              let coloredCell = entireSheet.getCell(k + startingRowIndexForTimes, (i * spacer) + spacerValue);
              coloredCell.getFormat().getFill().setColor(manuallyAssignedRoomColor);
            }
            else if (isManuallyAssignedAndUndersized) {
              let coloredCell = entireSheet.getCell(k + startingRowIndexForTimes, (i * spacer) + spacerValue);
              coloredCell.getFormat().getFill().setColor(manuallyAssignedUndersizedRoomColor);
            }
            if (hasConflict) {
              let coloredCell = entireSheet.getCell(k + startingRowIndexForTimes, (i * spacer) + spacerValue);
              coloredCell.getFormat().getFill().setColor(conflictColor);
            }
          }
        }
      }




      // 


      function setColorsFromColorBlocksBasedOnSpacing() {

        let blocksToBeColored = rooms[i].colorBlocks;
        blocksToBeColored.sort((a, b) => (a[0] < b[0]) ? -1 : 1); // sort them by row

        for (let j = 0; j < rooms[i].colorBlocks.length; j++) {
          let startRow = blocksToBeColored[j][0] + startingRowIndexForTimes;
          let startColumn = blocksToBeColored[j][1] + 1; // One because first column in Calendars sheet is the times e.g. 800AM
          let rowsInRange = blocksToBeColored[j][2];

          // This for loop looks backwards from where we are in the rooms.colorBlocks array and sees if we have switched 
          // to a new course / class; so if the unique ID of the course is different, we know to switch colors.
          // If we HAVEN'T switched, we stay with the same color. Otherwise we switch into a different color based on our colorIndex 
          // (see the switch statement for how we do this)
          // The rooms.colorBlocks are not stored in any particular way, so sometimes we'll have two blocks next to each other
          // that are on the same day of the week. We want to make sure those aren't 

          for (let k = j; k >= 0; k--) { // switch colors for new classes
            let currentDayOfWeek = blocksToBeColored[j][1];
            let dayOfWeekToCheckAgainst = blocksToBeColored[k][1]
            let currentClassRowID = blocksToBeColored[j][3];
            let classRowIDToCheckAgainst = blocksToBeColored[k][3];
            let currentBlockColor = blocksToBeColored[j][4];
            let blockColorToCheckAgainst = blocksToBeColored[k][4];
            if (currentDayOfWeek == dayOfWeekToCheckAgainst &&
              currentClassRowID != classRowIDToCheckAgainst) { // if it's in the same day of the week, but a different class ID

              switch (blockColorToCheckAgainst) {
                case 2:
                  if (currentBlockColor == 3) {
                    currentBlockColor = 3; // This stays the same because it is the undersized course color
                  }
                  else currentBlockColor = 0;
                  break;
                case 1:
                  if (currentBlockColor == 3) {
                    currentBlockColor = 3; // This stays the same because it is the undersized course color
                  }
                  else currentBlockColor = 2;
                  break;
                case 0:
                  if (currentBlockColor == 3) {
                    currentBlockColor = 3; // This stays the same because it is the undersized course color
                  }
                  else currentBlockColor = 1;
                  break;
                case 3:
                  if (currentBlockColor == 3) {
                    currentBlockColor = 3; // This stays the same because it is the undersized course color
                  }
                  else currentBlockColor = 0;
                  break;
              }

              blocksToBeColored[j][4] = currentBlockColor; // Here we set the color of the colorBlock based on the switch statement above
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

          // This last loop happens because sometimes we get to the end of the week and realize that we had to swap a class color.
          // In this event we go backwards and change all the color values for the previous days of the course
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
                classColor = mwfColors[colorIndex];// different colors for MWF TR
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




