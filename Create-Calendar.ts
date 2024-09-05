function main(workbook: ExcelScript.Workbook) {

  // Values for times and days
  let startingRowIndexForTimes = 2; // Which row is 800AM on in the chart? (It's actually row 3, but they are indexed from 0)
  let numberOfRowsInCalendarSheet: number = 53 // The row count for 8:00AM - 9:00PM. 

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
  let startingRowNumberForCourseData = startingRowIndexForCourseData + 1;
  let classData = courses.getRange("A" + startingRowNumberForCourseData as string + ":O" + rowCount as string).getValues();

  // Undersized classes
  let undersizedClassColor: string = "yellow"; // When courses get placed in rooms that are too large (based on the undersizedClassBuffer, below) we mark them this color
  let undersizedClassBuffer = .67; // Classes with enrollment that are less than this fraction of the room size are undersized

  // Coloring the chart
  let mwfColors: string[] = ["#BD9CDB", "#D07B82", "#8B7BD0", "#D07BB5", "#A77BD0", "#DB9CC7", "#D19CDB", "#A89CDB", "#DBB5B8", "#D3A6DB", "#C371DA"]; // The color blocks we use for MWF classes-- for ease of reading the chart
  let trColors: string[] = ["#B5DB9C", "#6AB39D", "#A6B36A", "#6AB383", "#87B36A", "#9CDBB2", "#9CDB9D", "#C5DC5C", "#9CDBC8", "#A6DBA7", "#72B8B8", "#ABDD09", "#60CBD5", "#D7D276"]; // same as above, but for TTh
  let colorIndex = 0; // A number we use to track which set of colors to use: mwf or tth
  let assignedRoomColor = "dbdbdb"; // Color for when we have "locked" or assigned a room (possibly redundant)
  let assignedAndUnderSizedColor = "red"; // Color for when it is "locked" or assigned, and undersized (possibly redundant)
  let manuallyAssignedRoomColor = "green" // Color for when we have "locked" or assigned a room (possibly redundant)
  let manuallyAssignedUndersizedRoomColor = "6161FE" // Color for when it is "locked" or assigned, and undersized (possibly redundant)
  let conflictColor = "red" // Color for when it is "locked" or assigned, and undersized (possibly redundant)

  // We store the class data and its color value (for populating the calendar and making it readable)
  type classAndColor = {
    classData: string;
    colorData: string;
    uniqueIndex: number;
  }


  // Start your engines!

  getDataFromCoursesSheet(); // Grab the data from the "Courses" sheet
  createRooms(); // Create the rooms data from the "Rooms" sheet
  createScheduleBasics(); // Create the schedule to be filled (this is just the data container, it has no data yet)
  fillInManuallyAssignedClasses(); // Fill in any classes we have already "locked" or assigned manually 
  MakeTempSchedule(); // Attempt to fill in courses into the schedule based on size, availability, etc.
  createRoomSchedule(); // Create the schedule for each room based on the above schedule

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

  // This creates the header for the rooms in the calendar sheet
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

  // This fills the arrays with empty values so we can put classes in specific indexes of the schedule
  // We build one of these for each room, using a Map with the five days of the week, and a 
  // number of rows equal to the rows in the calendar sheet (53) for 8:00AM - 9:00PM
  // This schedule will be filled in with class data prior to actually putting it on the sheet in the final step.
  function initializeRoomSchedule(): Map<string, classAndColor[]> {
    let schedule: Map<string, classAndColor[]> = new Map<string, classAndColor[]>(); // Where the first string is day of the week, and then the 
    //class data per cell and color data per cell
    let defaultValuesM: classAndColor[] = [];
    let defaultValuesT: classAndColor[] = [];
    let defaultValuesW: classAndColor[] = [];
    let defaultValuesR: classAndColor[] = [];
    let defaultValuesF: classAndColor[] = [];

    for (let i = 0; i < numberOfRowsInCalendarSheet; i++) {
      defaultValuesM.push({ classData: "" as string, colorData: "" as string, uniqueIndex: -1 as number });
      defaultValuesT.push({ classData: "" as string, colorData: "" as string, uniqueIndex: -1 as number });
      defaultValuesW.push({ classData: "" as string, colorData: "" as string, uniqueIndex: -1 as number });
      defaultValuesR.push({ classData: "" as string, colorData: "" as string, uniqueIndex: -1 as number });
      defaultValuesF.push({ classData: "" as string, colorData: "" as string, uniqueIndex: -1 as number });
    }

    schedule.set("M", defaultValuesM);
    schedule.set("T", defaultValuesT);
    schedule.set("W", defaultValuesW);
    schedule.set("R", defaultValuesR);
    schedule.set("F", defaultValuesF);

    return schedule;
  }

  // Fill classes that have aready been assigned
  function fillInManuallyAssignedClasses() {
    let timeValues = calendar.getRange("A3:A56").getValues(); // The time values we are working with, generally from 8:00AM to 9:00PM
    for (let i = 0; i < uvaClasses.length; i++) {
      // If the room has been assigned, we find that room in our rooms array and the start time in our timeValues array, 
      // then fill the slot in our rooms.schedule array, noting the special colors for "locked" or assigned rooms
      if (uvaClasses[i].assignedRoom != "") {
        // Find which room it has been assigned to
        let roomIndex = rooms.findIndex(room => room.name == uvaClasses[i].assignedRoom);

        // Find the index of the timeValues that coresponds to when the class starts
        let timeIndex = timeValues.findIndex(time => time[0] == uvaClasses[i].startTime);

        let startRow = timeIndex;
        let courseDuration = 0;

        // Find the course duration
        for (let j = timeIndex; j < timeValues.length; j++) {
          if (uvaClasses[i].endTime > timeValues[j][0]) {
            courseDuration++;
          }
          else break;
        }

        // Check if it is under capacity. If so, fill the slot with the color for an under-capacity room.
        // If not, fill it with the manually-assigned color.
        if (uvaClasses[i].enrollment < rooms[roomIndex].capacity * undersizedClassBuffer) {
          fillOpenSlot(uvaClasses[i], startRow, courseDuration, roomIndex, "undersizedColor");
          let courseRowNumber = uvaClasses[i].rowInDatabase;
          reportAssignment(courseRowNumber, "assigned manually and undersized");
        }
        else {
          fillOpenSlot(uvaClasses[i], startRow, courseDuration, roomIndex, "manuallyAssignedColor");
          let courseRowNumber = uvaClasses[i].rowInDatabase;
          reportAssignment(courseRowNumber, "assigned manually");
        }
      }
    }
  }

  function MakeTempSchedule() {

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
          let wholeCourseDurationOpenOnCalendar = 1; // boolean as number; we check every day and time and then multiply all these values together-- if, for instance, in a MWF class there is a zero for "F" it won't place the course in MW
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
            if (uvaClasses[i].enrollment < rooms[roomFoundIndex].capacity * undersizedClassBuffer) {
              fillOpenSlot(uvaClasses[i], currentRow, courseDuration, roomFoundIndex, "undersizedColor");
            }
            else { 
              fillOpenSlot(uvaClasses[i], currentRow, courseDuration, roomFoundIndex, "automaticColors");
            }
            let rowNumber = uvaClasses[i].rowInDatabase;
            if (roomFoundIndex < 16) { // 16 is the index where "Unassigned" begins in the rooms list array
              reportAssignment(rowNumber, "automatically assigned");
            }
            else {
              reportAssignment(rowNumber, "unassigned");
            }
          }
          else {
            let rowNumber = uvaClasses[i].rowInDatabase;
            reportAssignment(rowNumber, "no");
            
          }
        }
      }
    }
  }

  function reportAssignment(rowNumber: number, value: string) {
    courses.getCell(rowNumber, 15).setValue(value);
    if (value.includes("already") || value == "no") {
      courses.getCell(rowNumber, 15).getFormat().getFill().setColor("red");
    }
    else if (value == "unassigned") {
      courses.getCell(rowNumber, 15).getFormat().getFill().setColor("pink");
    }
    else {
      courses.getCell(rowNumber, 15).getFormat().getFill().setColor("#cccccc");
    }


  }


  // This attempts to find an open slot in all our rooms.schedule arrays within the rooms array
  function attemptToFindOpenSlot(uvaClass: UVAClass, row: number, courseDuration: number): [number, number] {

    let foundSpot = 0; // using number as boolean since we check all days; if MW are 1 and F is zero, then MWF is zero since they are multiplied at the end; 
    // this way we can check all days to verify the whole schedule is open
    var roomIndex: number = 0;
    let daysOfWeek: number[] = []; // we capture which days to populate on the sheet; this array let's us know which days we're dealing with for this course

    for (let i = 0; i < rooms.length; i++) { // Go  through all rooms, which have been ordered by size, and check for open slots
      let truthValues: number[] = []; // We collect a bunch of values and then multiply them together at the end-- if a single zero (0) is in the lot, signifying one of the days is full, it is false
      if (rooms[i].capacity >= uvaClass.enrollment) {

        foundSpot = 1; // we have a room that could hypothetically hold the class, but we need to check each day

        // The rooms array contains rooms.schedule arrays that then contain dictionaries with key values for each day "M", "T", "W", etc.
        // Within those are arrays for the slots; importantly, the rooms.schedules don't know or care about times, they only care about slots; 
        // times are handled in MakeTempSchedule function and converted to index values to check slots in these arrays
        for (let j = 0; j < courseDuration; j++) { // we need to check the whole duration; again, duration is slots, not actual times, which this array doesn't 
          // know nor care about
          if (uvaClass.day.includes("M")) {
            if (rooms[i].schedule.get("M")[row + j].classData == "") {
              truthValues.push(1);
              if (!daysOfWeek.includes(0)) // If the array doesn't already contain 0 (signifying Monday as the 0 column in the chart), put in 0
                daysOfWeek.push(0);
            }
            else truthValues.push(0);
          }
          if (uvaClass.day.includes("T")) {

            if (rooms[i].schedule.get("T")[row + j].classData == "") {
              truthValues.push(1);
              if (!daysOfWeek.includes(1))
                daysOfWeek.push(1);
            }
            else truthValues.push(0);
          }
          if (uvaClass.day.includes("W")) {
            if (rooms[i].schedule.get("W")[row + j].classData == "") {
              truthValues.push(1);
              if (!daysOfWeek.includes(2))
                daysOfWeek.push(2);
            }
            else truthValues.push(0);
          }
          if (uvaClass.day.includes("R")) {
            if (rooms[i].schedule.get("R")[row + j].classData == "") {
              truthValues.push(1);
              if (!daysOfWeek.includes(3))
                daysOfWeek.push(3);
            }
            else truthValues.push(0);
          }
          if (uvaClass.day.includes("F")) {
            if (rooms[i].schedule.get("F")[row + j].classData == "") {
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
        break;
      }

    }

    return [foundSpot, roomIndex];

  }


  function fillOpenSlot(uvaClass: UVAClass, startRow: number, courseDuration: number, foundRoomIndex: number, roomColorInfo: string) {
    let actualRowNumber = uvaClass.rowInDatabase + 1; // Zero index nonsense
    let courseInfo = uvaClass.courseMnemonic + " " + uvaClass.courseNumber + " " + uvaClass.courseSection + " [" + uvaClass.enrollment + "] " + " <" + actualRowNumber + ">" + " {" + uvaClass.uniqueIndex + "}";

    let classDaysOfWeek = "";

    checkDaysOfWeek(checkIfManuallyAssigned);
    checkDaysOfWeek(assignManualColorValues);
    checkDaysOfWeek(assignScheduleValues);

    function checkDaysOfWeek(fn: (dayOfWeek: string) => void) {
      if (uvaClass.day.includes("M")) {
        classDaysOfWeek = "MWF"
        fn("M")
      }
      if (uvaClass.day.includes("T")) {
        classDaysOfWeek = "TR"
        fn("T")
      }
      if (uvaClass.day.includes("W")) {
        fn("W")
      }
      if (uvaClass.day.includes("R")) {
        fn("R")
      }
      if (uvaClass.day.includes("F")) {
        fn("F")
      }
    }

    function checkIfManuallyAssigned(dayOfWeek: string) {// If we assigned two or more classes to the same room at the same time
      // This can happen due to manual assignment errors; it won't happen automatically
      for (let row = startRow; row < startRow + courseDuration; row++) {
        if (rooms[foundRoomIndex].schedule.get(dayOfWeek)[row].classData != "") {
          let currentInfoInRoom = rooms[foundRoomIndex].schedule.get(dayOfWeek)[row].classData;
          courseInfo = currentInfoInRoom + " /-/ " + courseInfo;
          roomColorInfo = "double-booked";
          console.log("ERROR: duplicate manual assignment for " + uvaClass.courseMnemonic + " " + uvaClass.courseNumber);

          let courseRowNumber = uvaClass.rowInDatabase;
          reportAssignment(courseRowNumber, "room already manually assigned to " + currentInfoInRoom);
        }
      }
    }

    function assignScheduleValues(dayOfWeek: string) {
      for (let row = startRow; row < startRow + courseDuration; row++) {
        rooms[foundRoomIndex].schedule.get(dayOfWeek)[row].classData = courseInfo;
        rooms[foundRoomIndex].schedule.get(dayOfWeek)[row].uniqueIndex = uvaClass.uniqueIndex;
      }
    }

    function assignManualColorValues(dayOfWeek: string) {
      // If they are manually assigned or undersized, we manually assign color values
      for (let i = startRow; i < startRow + courseDuration; i++) {
        if (roomColorInfo == "manuallyAssignedColor") {
          rooms[foundRoomIndex].schedule.get(dayOfWeek)[i].colorData = manuallyAssignedRoomColor;
        }
        else if (roomColorInfo == "undersizedColor") {
          rooms[foundRoomIndex].schedule.get(dayOfWeek)[i].colorData = undersizedClassColor;
        }
        else if (roomColorInfo == "double-booked") {
          rooms[foundRoomIndex].schedule.get(dayOfWeek)[i].colorData = conflictColor;
          
        }

      }

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
      console.log("filling room: " + rooms[i].name);
      let roomNameCell = calendar.getCell(0, (i * spacer) + 1); // Get every 7th cell, staring in the 2nd column

      roomNameCell.setValue(rooms[i].name);
      roomNameCell.getFormat().getFont().setBold(true);
      let uniqueIndexesAndColorIndexes: number[][] = []
      let availableMWFColorIndexes: number[] = []
      let availableTRColorIndexes: number[] = []

      // Populate with the index values of the colors (which both have the same length)
      for (let j = 0; j < mwfColors.length; j++) {
        availableMWFColorIndexes.push(j);
      }
      for (let j = 0; j < trColors.length; j++) {
        availableTRColorIndexes.push(j);
      }

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
          let classData = rooms[i].schedule.get(dayChar)[j].classData;
          if (classData == "") {
            classData = "-";
          }
          courseScheduleData[j][0] = classData;
        }
        let dayCell = entireSheet.getCell(1, (i * spacer) + spacerValue);
        dayCell.setValue(dayOfWeek);
        let targetCell = entireSheet.getCell(2, (i * spacer) + spacerValue);
        let targetRange = targetCell.getResizedRange(rooms[i].schedule.get(dayChar).length - 1, 0);

        // Set actual data (we do this as a block)
        targetRange.setValues(courseScheduleData);

        // Set color data for cells in the generated sheet  
        for (let j = 0; j < courseScheduleData.length; j++) {

          let uniqueIndex: number = rooms[i].schedule.get(dayChar)[j].uniqueIndex;
          let colorData = rooms[i].schedule.get(dayChar)[j].colorData;
          let classData = rooms[i].schedule.get(dayChar)[j].classData;

          // Check if we manually assigned color data due to it being undersized, etc.
          if (colorData != "") {
            
          }

          // First we check if the class has already been assigned a color this week
          else if (uniqueIndexesAndColorIndexes.some(subArray => subArray[0] == uniqueIndex)) {
            let subArrayIndex = uniqueIndexesAndColorIndexes.findIndex(subArray => subArray[0] == uniqueIndex);
            let subArrayColorIndex = uniqueIndexesAndColorIndexes[subArrayIndex][1];
            if (dayChar == "M" || dayChar == "W" || dayChar == "F") {
              rooms[i].schedule.get(dayChar)[j].colorData = mwfColors[subArrayColorIndex];
            }
            else {
              rooms[i].schedule.get(dayChar)[j].colorData = trColors[subArrayColorIndex];
            }
          }

          // Then if it hasn't already been added to the list, we check to see if it is different from the one above it
          else if (classData != "") {
            if (dayChar == "M" || dayChar == "W" || dayChar == "F") {
              rooms[i].schedule.get(dayChar)[j].colorData = mwfColors[availableMWFColorIndexes[0]];
              uniqueIndexesAndColorIndexes.push([uniqueIndex, availableMWFColorIndexes[0]]);
              availableMWFColorIndexes.splice(0, 1);

            }
            else {
              rooms[i].schedule.get(dayChar)[j].colorData = trColors[availableTRColorIndexes[0]];
              uniqueIndexesAndColorIndexes.push([uniqueIndex, availableTRColorIndexes[0]]);
              availableTRColorIndexes.splice(0, 1);
            }
          }


        }

        // Finally we actually add the colors
        for (let j = 0; j < courseScheduleData.length; j++) {
          let coloredCell = entireSheet.getCell(j + startingRowIndexForTimes, (i * spacer) + spacerValue);
          let colorData = rooms[i].schedule.get(dayChar)[j].colorData;

          if (colorData != "") {
            coloredCell.getFormat().getFill().setColor(colorData);
          }
        }
      }


    }

  }

}




