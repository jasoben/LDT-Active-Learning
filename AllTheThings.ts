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
  
  let rooms: room[] = []

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

  
    
  var courses = workbook.getWorksheet("Courses");
  var calendar = workbook.getWorksheet("Calendar");
  
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
  let uvaClasses: UVAClass[] = [];
  let classData = courses.getRange("A4:O" + rowCount as string).getValues();

  for (let i = 0; i < classData.length; i++) {
    uvaClasses[i] = {
        courseMnemonic: classData[i][0] as string,
        courseNumber: classData[i][1] as string,
        day: classData[i][5] as string,
        enrollment: classData[i][9] as number,
        startTime: classData[i][6] as number,
        endTime: classData[i][7] as number,
        assignedRoom: classData[i][12] as string,
        uniqueIndex: classData[i][14] as number
    };
  }


  function InitializeRoomSchedule() : Map<string, string[]> {
    let schedule: Map<string, string[]> = new Map<string, string[]>();
    let defaultTimes: string[] = [];
    
    for (let i = 0; i < 53; i++) 
    {
        defaultTimes[i] = "";
    }
    
    schedule.set("M", defaultTimes);
    schedule.set("T", defaultTimes);
    schedule.set("w", defaultTimes);
    schedule.set("Th", defaultTimes);
    schedule.set("F", defaultTimes);

    return schedule;
  }
    
}