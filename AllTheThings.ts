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
      schedule: Map<string, string[]>; // e.g. schedule["M"] = monday's schedule (which is times that are filled by class names)
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
  
  var entireSheet: ExcelScript.Range = calendar.getRange("A1:Z1000");

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
          uniqueIndex: classData[i][14] as number
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
    
  function InitializeRoomSchedule() : Map<string, string[]> {
    let schedule: Map<string, string[]> = new Map<string, string[]>();
    let defaultTimesM: string[] = [];
    let defaultTimesTu: string[] = [];
    let defaultTimesW: string[] = [];
    let defaultTimesTh: string[] = [];
    let defaultTimesF: string[] = [];
    
    for (let i = 0; i < 53; i++) 
    {
      defaultTimesM[i] = "";
      defaultTimesTu[i] = "";
      defaultTimesW[i] = "";
      defaultTimesTh[i] = "";
      defaultTimesF[i] = "";
    }
    
    schedule.set("M", defaultTimesM);
    schedule.set("T", defaultTimesTu);
    schedule.set("W", defaultTimesW);
    schedule.set("Th", defaultTimesTh);
    schedule.set("F", defaultTimesF);

    
    return schedule;
  }

  function AttemptToMakeSchedule() {
    CreateScheduleBasics();
    let times = calendar.getRange("A3:A56");
    let rowCount = times.getRowCount();
    let cellValues = times.getValues() as number[][];
    for (let i = 0; i < rowCount; i++)
    {
      for (let j = 0; j < uvaClasses.length; j++)
      {
        if (cellValues[i][0] == uvaClasses[j].startTime) {
          AttemptToFindOpenSlot(uvaClasses[j], i);
        }
      }
    }
    
    CreateRoomSchedule();
  }

  function AttemptToFindOpenSlot(uvaClass:UVAClass, row:number) {
    let courseInfo = uvaClass.courseMnemonic + " " + uvaClass.courseNumber + " " + uvaClass.courseSection;

    let foundSpot = false;
    for (let i = 0; i < rooms.length; i++) {
      if (rooms[i].capacity >= uvaClass.enrollment) {
        if (uvaClass.day.includes("M")) {
          if (rooms[i].schedule.get("M")[row] == "") {
            rooms[i].schedule.get("M")[row] = courseInfo;
            foundSpot = true;
          }
          else {
            return;
          }
        }
        if (uvaClass.day == "TuTh" || uvaClass.day == "T") {
          if (rooms[i].schedule.get("T")[row] == "") {
            rooms[i].schedule.get("T")[row] = courseInfo;
            foundSpot = true;
            
          }
          else {
            return;
          }
        }
        if (uvaClass.day.includes("W")) {
          if (rooms[i].schedule.get("W")[row] == "") {
            rooms[i].schedule.get("W")[row] = courseInfo;
            foundSpot = true;
          }
          else {
            return;
          }
        }
        if (uvaClass.day.includes("Th")) {
          if (rooms[i].schedule.get("Th")[row] == "") {
            rooms[i].schedule.get("Th")[row] = courseInfo;
            foundSpot = true;
          }
          else {
            return;
          }
        }
        if (uvaClass.day.includes("F")) {
          if (rooms[i].schedule.get("F")[row] == "") {
            rooms[i].schedule.get("F")[row] = courseInfo;
            foundSpot = true;
          }
          else {
            return;
          }
        }

        
      }
      if (foundSpot)
        break;

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
      for (let j = 0; j < rooms[i].schedule.get("M").length; j++) {
        entireSheet.getCell(j + 2, (i * spacer) + 1).setValue(rooms[i].schedule.get("M")[j]);
        
      }
      
      let tuesdayCell = entireSheet.getCell(1, (i * spacer) + 2);
      tuesdayCell.setValue("Tuesday");
      for (let j = 0; j < rooms[i].schedule.get("T").length; j++) {
        entireSheet.getCell(j + 2, (i * spacer) + 2).setValue(rooms[i].schedule.get("T")[j]);
      }

      let wednesdayCell = entireSheet.getCell(1, (i * spacer) + 3);
      wednesdayCell.setValue("Wednesday");
      for (let j = 0; j < rooms[i].schedule.get("W").length; j++) {
        entireSheet.getCell(j + 2, (i * spacer) + 3).setValue(rooms[i].schedule.get("W")[j]);
      }

      let thursdayCell = entireSheet.getCell(1, (i * spacer) + 4);
      thursdayCell.setValue("Thursday");
      for (let j = 0; j < rooms[i].schedule.get("Th").length; j++) {
        entireSheet.getCell(j + 2, (i * spacer) + 4).setValue(rooms[i].schedule.get("Th")[j]);
      }

      let fridayCell = entireSheet.getCell(1, (i * spacer) + 5);
      fridayCell.setValue("Friday");
      for (let j = 0; j < rooms[i].schedule.get("F").length; j++) {
        entireSheet.getCell(j + 2, (i * spacer) + 5).setValue(rooms[i].schedule.get("F")[j]);
      }
    }
  }
}

    
