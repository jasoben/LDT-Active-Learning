function main(workbook: ExcelScript.Workbook) {

  let cellHasValue: number[][] = [];
  let cellValue: string[] = [];
  let colors: string[] = ["b4f2ac", "dfe8ae", "b2bded", "adf0dd"]
  let colorIndex = 0;
  let roomHeadings: string[][] = [];
  let days: string[] = ["M", "T", "W", "Th", "F"]
  let startOfClass: boolean = false;
  var classRooms18to30: ExcelScript.Worksheet;
  var classRooms31to54: ExcelScript.Worksheet;
  var classRooms55to62: ExcelScript.Worksheet;
  var classRooms63to99: ExcelScript.Worksheet;
  var classRoomsTimes: number[][] = [];
  let startNewCourseBlock: boolean = false;
  
  type room = {
      name: string;
      capacity: number;
      schedule: Map<string, [string[], string[]]>; // e.g. schedule["M"] = monday's schedule (which is times that are filled by class names). The two strings of the tuple are the schedule cells themselves (to be filled with data) and the cell color values.

  }

  var NCH107: room, 
    NCH191: room,
    WIL244: room,
    CHEM204: room,
    CHEM206: room,
    UNASSIGNEDROOM : room;
  
  let rooms: room[] = []

  var courses = workbook.getWorksheet("Courses");
  var calendar = workbook.getWorksheet("Calendar");
  
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

  AssignRooms();
  GetDataFromCoursesSheet();
  AttemptToMakeSchedule();

  function GetDataFromCoursesSheet() {
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
          rowInDatabase: i + 4 // the data starts on row 4
      };
    }
  }

  function AssignRooms() {
      // Your code here
    let UNASSIGNEDROOM: room = {
      name: "Unassigned",
      capacity: 0,
      schedule: InitializeRoomSchedule()
    }
    
    let NCH107: room = {
        name: "NCH 107",
        capacity: 30,
      schedule: InitializeRoomSchedule()
    }

    rooms.push(NCH107);

    let NCH191: room = {
      name: "NCH 191",
      capacity: 30,
      schedule: InitializeRoomSchedule()
    }

    rooms.push(NCH191);

    let WIL244: room = {
      name: "Wilson 244",
      capacity: 30,
      schedule: InitializeRoomSchedule()
    }

    rooms.push(WIL244);

    let CHEM206: room = {
      name: "CHEM 206",
      capacity: 54,
      schedule: InitializeRoomSchedule()
    }

    rooms.push(CHEM206);

    let CHEM204: room = {
      name: "CHEM 204",
      capacity: 54,
      schedule: InitializeRoomSchedule()
    }

    rooms.push(CHEM204);
          
    rooms.sort( (a,b) => (a.capacity < b.capacity) ? -1 : 1);
  }

  function InitializeRoomSchedule(): Map<string, [string[], string[]]> {
    let schedule: Map<string, [string[], string[]]> = new Map<string, [string[], string[]]>();
    let defaultTimesM: string[] = [];
    let defaultTimesTu: string[] = [];
    let defaultTimesW: string[] = [];
    let defaultTimesTh: string[] = [];
    let defaultTimesF: string[] = [];

    let defaultColorsM: string[] = [];
    let defaultColorsTu: string[] = [];
    let defaultColorsW: string[] = [];
    let defaultColorsTh: string[] = [];
    let defaultColorsF: string[] = [];

    for (let i = 0; i < 53; i++) {
      defaultTimesM[i] = "";
      defaultTimesTu[i] = "";
      defaultTimesW[i] = "";
      defaultTimesTh[i] = "";
      defaultTimesF[i] = "";

      defaultColorsM[i] = "";
      defaultColorsTu[i] = "";
      defaultColorsW[i] = "";
      defaultColorsTh[i] = "";
      defaultColorsF[i] = "";
    }

    schedule.set("M", [defaultTimesM, defaultColorsM]);
    schedule.set("T", [defaultTimesTu, defaultColorsTu]);
    schedule.set("W", [defaultTimesW, defaultColorsW]);
    schedule.set("Th", [defaultTimesTh, defaultColorsTh]);
    schedule.set("F", [defaultTimesF, defaultColorsF]);

    return schedule;
  }

  function AttemptToMakeSchedule() {
    CreateScheduleBasics();
    let times = calendar.getRange("A3:A56");
    let rowCount = times.getRowCount();
    let cellTimeValues = times.getValues() as number[][];
    for (let i = 0; i < uvaClasses.length; i++) {
      for (let j = 0; j < rowCount; j++) {
        if (cellTimeValues[j][0] >= uvaClasses[i].startTime && cellTimeValues[j][0] < uvaClasses[i].endTime) {
          let wholeCourseDurationOpenOnCalendar = 1; // boolean as number
          var foundRoomIndex : number;
       
          for (let k = 0; k < 10; k++) { // we need to check for the whole duration of the course
            if (cellTimeValues[j + k][0] >= uvaClasses[i].startTime && cellTimeValues[j + k][0] < uvaClasses[i].endTime) {
              let attemptInfo = AttemptToFindOpenSlot(uvaClasses[i], j + k);
              wholeCourseDurationOpenOnCalendar = attemptInfo[0] * wholeCourseDurationOpenOnCalendar;
              foundRoomIndex = attemptInfo[1];
            }
            else 
              break;
          }
          if (wholeCourseDurationOpenOnCalendar == 1) {
            FillOpenSlot(uvaClasses[i], j, foundRoomIndex);
          }
        }
      }

      colorIndex + 1 < colors.length ? colorIndex++ : colorIndex = 0; // cycle to a new color for the next class    

    }
    
    CreateRoomSchedule();
  }
  

  function AttemptToFindOpenSlot(uvaClass: UVAClass, row: number) : [number, number] {

    let foundSpot = 0; // using number as boolean
    var roomIndex : number;

    for (let i = 0; i < rooms.length; i++) {
      let truthValues: number[] = [];
      if (rooms[i].capacity >= uvaClass.enrollment) {

        foundSpot = 1; // we have a room that could hypothetically hold the class, but we need to check each day

        if (uvaClass.day.includes("M")) {
          if (rooms[i].schedule.get("M")[0][row] == "") {
            truthValues.push(1);
          }
          else truthValues.push(0);
        }
        if (uvaClass.day == "TuTh" || uvaClass.day == "T") {
          
          if (rooms[i].schedule.get("T")[0][row] == "") {
            truthValues.push(1);
          }
          else truthValues.push(0);
        }
        if (uvaClass.day.includes("W")) {
          if (rooms[i].schedule.get("W")[0][row] == "") {
            truthValues.push(1);
          }
          else truthValues.push(0);
        }
        if (uvaClass.day.includes("Th")) {
          if (rooms[i].schedule.get("Th")[0][row] == "") {
            truthValues.push(1);
          }
          else truthValues.push(0);
        }
        if (uvaClass.day.includes("F")) {
          if (rooms[i].schedule.get("F")[0][row] == "") {
            truthValues.push(1);
          }
          else truthValues.push(0);
        }
      }

      for (let j = 0; j < truthValues.length; j++) {
        foundSpot = foundSpot * truthValues[j];
      }

      if (foundSpot == 1) {
        roomIndex = i;
        break; 
      }
      
    }
    
    
    return [ foundSpot, roomIndex];
  }

  function FillOpenSlot(uvaClass: UVAClass, row: number, foundRoomIndex: number) {
    let courseInfo = uvaClass.courseMnemonic + " " + uvaClass.courseNumber + " " + uvaClass.courseSection + " " + uvaClass.rowInDatabase;

    if (uvaClass.day.includes("M")) {
      rooms[foundRoomIndex].schedule.get("M")[0][row] = courseInfo;
      rooms[foundRoomIndex].schedule.get("M")[1][row] = colors[colorIndex];
    }
    if (uvaClass.day == "TuTh" || uvaClass.day == "T") {
      rooms[foundRoomIndex].schedule.get("T")[0][row] = courseInfo;
      rooms[foundRoomIndex].schedule.get("T")[1][row] = colors[colorIndex];
    }
    if (uvaClass.day.includes("W")) {
      rooms[foundRoomIndex].schedule.get("W")[0][row] = courseInfo;
      rooms[foundRoomIndex].schedule.get("W")[1][row] = colors[colorIndex];
    }
    if (uvaClass.day.includes("Th")) {
      rooms[foundRoomIndex].schedule.get("Th")[0][row] = courseInfo;
      rooms[foundRoomIndex].schedule.get("Th")[1][row] = colors[colorIndex];
    }
    if (uvaClass.day.includes("F")) {
      rooms[foundRoomIndex].schedule.get("F")[0][row] = courseInfo;
      rooms[foundRoomIndex].schedule.get("F")[1][row] = colors[colorIndex];
    }

  }

  function CreateScheduleBasics() {
    entireSheet.clear();
    let hour = 8;
    let mod = 0;
    let minutes = ["00", "15", "30", "45"];
    
    
    for (let i = 2; i < 56; i++)
    {
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

  function CreateRoomSchedule() {

    let spacer = 6;

    for (let i = 0; i < rooms.length; i++)
    {
      let roomNameCell = entireSheet.getCell(0, (i * spacer) + 1); // Get every 7th cell, staring in the 2nd column
      roomNameCell.setValue(rooms[i].name);

      let mondayCell = entireSheet.getCell(1, (i * spacer) + 1);
      mondayCell.setValue("Monday");
      for (let j = 0; j < rooms[i].schedule.get("M")[0].length; j++) {
        entireSheet.getCell(j + 2, (i * spacer) + 1).setValue(rooms[i].schedule.get("M")[0][j]);
        if (rooms[i].schedule.get("M")[1][j] != "")
          entireSheet.getCell(j + 2, (i * spacer) + 1).getFormat().getFill().setColor(rooms[i].schedule.get("M")[1][j]);
      }
      
      let tuesdayCell = entireSheet.getCell(1, (i * spacer) + 2);
      tuesdayCell.setValue("Tuesday");
      for (let j = 0; j < rooms[i].schedule.get("T")[0].length; j++) {
        entireSheet.getCell(j + 2, (i * spacer) + 2).setValue(rooms[i].schedule.get("T")[0][j]);
        if (rooms[i].schedule.get("T")[1][j] != "")
          entireSheet.getCell(j + 2, (i * spacer) + 2).getFormat().getFill().setColor(rooms[i].schedule.get("T")[1][j]);
      }

      let wednesdayCell = entireSheet.getCell(1, (i * spacer) + 3);
      wednesdayCell.setValue("Wednesday");
      for (let j = 0; j < rooms[i].schedule.get("W")[0].length; j++) {
        entireSheet.getCell(j + 2, (i * spacer) + 3).setValue(rooms[i].schedule.get("W")[0][j]);
        if (rooms[i].schedule.get("W")[1][j] != "")
          entireSheet.getCell(j + 2, (i * spacer) + 3).getFormat().getFill().setColor(rooms[i].schedule.get("W")[1][j]);
      }

      let thursdayCell = entireSheet.getCell(1, (i * spacer) + 4);
      thursdayCell.setValue("Thursday");
      for (let j = 0; j < rooms[i].schedule.get("Th")[0].length; j++) {
        entireSheet.getCell(j + 2, (i * spacer) + 4).setValue(rooms[i].schedule.get("Th")[0][j]);
        if (rooms[i].schedule.get("Th")[1][j] != "")
          entireSheet.getCell(j + 2, (i * spacer) + 4).getFormat().getFill().setColor(rooms[i].schedule.get("Th")[1][j]);
      }

      let fridayCell = entireSheet.getCell(1, (i * spacer) + 5);
      fridayCell.setValue("Friday");
      for (let j = 0; j < rooms[i].schedule.get("F")[0].length; j++) {
        entireSheet.getCell(j + 2, (i * spacer) + 5).setValue(rooms[i].schedule.get("F")[0][j]);
        if (rooms[i].schedule.get("F")[1][j] != "")
          entireSheet.getCell(j + 2, (i * spacer) + 5).getFormat().getFill().setColor(rooms[i].schedule.get("F")[1][j]);
      }
    }
  }
}

    
