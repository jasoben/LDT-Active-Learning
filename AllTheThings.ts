function main(workbook: ExcelScript.Workbook) {

  let cellHasValue: number[][] = [];
  let cellValue: string[] = [];
  let colors: string[] = ["b4f2ac", "dfe8ae", "b2bded", "adf0dd"]
  let colorIndex = 0;
  let roomHeadings: string[][] = [];
  let days: string[] = ["M", "T", "W", "Th", "F"]
  let startOfClass: boolean = false;
  var classRoomsTimes: number[][] = [];
  let startNewCourseBlock: boolean = false;
  
  type room = {
      name: string;
      capacity: number;
      schedule: Map<string, string[][]>; // e.g. schedule["M"] = monday's schedule (which is times that are filled by class names). The two strings of the tuple are the schedule cells themselves (to be filled with data) and the cell color values.

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
    
    let roomsRange = roomsSheet.getUsedRange();
    let roomsCount = roomsRange.getRowCount();
    let roomsValues = roomsRange.getValues();
    for (let i = 1; i < roomsCount; i++) { // 1 because first row is heading
      let thisRoom: room = {
        name: roomsValues[i][0] as string, 
        capacity: roomsValues[i][1] as number,
        schedule: InitializeRoomSchedule()
      }

      rooms.push(thisRoom);
    }
    rooms.sort( (a,b) => (a.capacity < b.capacity) ? -1 : 1);
  }

  function InitializeRoomSchedule(): Map<string, string[][]> {
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
          if (rooms[i].schedule.get("M")[row][0] == "") {
            truthValues.push(1);
          }
          else truthValues.push(0);
        }
        if (uvaClass.day == "TuTh" || uvaClass.day == "T") {
          
          if (rooms[i].schedule.get("T")[row][0] == "") {
            truthValues.push(1);
          }
          else truthValues.push(0);
        }
        if (uvaClass.day.includes("W")) {
          if (rooms[i].schedule.get("W")[row][0] == "") {
            truthValues.push(1);
          }
          else truthValues.push(0);
        }
        if (uvaClass.day.includes("Th")) {
          if (rooms[i].schedule.get("Th")[row][0] == "") {
            truthValues.push(1);
          }
          else truthValues.push(0);
        }
        if (uvaClass.day.includes("F")) {
          if (rooms[i].schedule.get("F")[row][0] == "") {
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
      rooms[foundRoomIndex].schedule.get("M")[row][0] = courseInfo;
    }
    if (uvaClass.day == "TuTh" || uvaClass.day == "T") {
      rooms[foundRoomIndex].schedule.get("T")[row][0] = courseInfo;
    }
    if (uvaClass.day.includes("W")) {
      rooms[foundRoomIndex].schedule.get("W")[row][0] = courseInfo;
    }
    if (uvaClass.day.includes("Th")) {
      rooms[foundRoomIndex].schedule.get("Th")[row][0] = courseInfo;
    }
    if (uvaClass.day.includes("F")) {
      rooms[foundRoomIndex].schedule.get("F")[row][0] = courseInfo;
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
    
    for (let i = 0; i < rooms.length; i++) {
      let roomNameCell = entireSheet.getCell(0, (i * spacer) + 1); // Get every 7th cell, staring in the 2nd column
      roomNameCell.setValue(rooms[i].name);

      function makeColumn(spacerValue:number, dayOfWeek: string, dayChar: string) {
        let dayCell = entireSheet.getCell(1, (i * spacer) + spacerValue);
        dayCell.setValue(dayOfWeek);
        let targetCell = entireSheet.getCell(2, (i * spacer) + spacerValue);
        let targetRange = targetCell.getResizedRange(rooms[i].schedule.get(dayChar).length - 1, 0);
        
        targetRange.setValues(rooms[i].schedule.get(dayChar));
      }

      makeColumn(1, "Monday", "M");
      makeColumn(2, "Tuesday", "T");
      makeColumn(3, "Wednesday", "W");
      makeColumn(4, "Thursday", "Th");
      makeColumn(5, "Friday", "F");

    } 
  }
}

    
