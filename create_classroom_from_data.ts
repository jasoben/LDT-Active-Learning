let cellHasValue: number[][] = [];
let cellValue: string[] = [];
let colors: string[] = ["b4f2ac", "dfe8ae", "b2bded", "adf0dd"]
let colorIndex = 0;
let roomHeadings: string[][] = [];
let days: string[] = ["M", "T", "W", "Th", "F"]
let startOfClass: boolean = false;
var classRooms: ExcelScript.Worksheet;
var classRoomsTimes: number[][] = [];
let startNewCourseBlock: boolean = false;

function main(workbook: ExcelScript.Workbook) {
    // Your code here

    var courses = workbook.getWorksheet("Courses");

    classRooms = workbook.getActiveWorksheet();

    if (classRooms.getName() == "Courses") {
        throw "don't run this on the courses page"
    }

    var classRoomRange = classRooms.getRange("B3:Z55");
    classRoomRange.clear();

    var enrollMin: number = classRooms.getRange("B57").getValue() as number;
    var enrollMax: number = classRooms.getRange("C57").getValue() as number;

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

    classRoomsTimes = classRooms.getRange("A3:A55").getValues() as number[][];

    roomHeadings = classRooms.getRange("B1:BZ1").getValues() as string[][];

    // Check classes for enrollment and assign to columns in sheet
    for (let i = 0; i < uvaClasses.length; i++) {
        if (uvaClasses[i].enrollment <= enrollMax && uvaClasses[i].enrollment >= enrollMin) {
            CheckClassRoomsTimes(uvaClasses[i].startTime, uvaClasses[i].endTime, uvaClasses[i].courseMnemonic, uvaClasses[i].courseNumber, uvaClasses[i].enrollment, uvaClasses[i].uniqueIndex, uvaClasses[i].assignedRoom, uvaClasses[i].day);
            colorIndex++; // We use this to change the fill colors on different classes
            if (colorIndex > colors.length - 1) {
                colorIndex = 0;
            }

        }
    }

}
function CheckClassRoomsTimes(startTime: number, endTime: number, courseMnemonic: string, courseNumber: string, enrollment: number, uniqueIndex: number, assignedRoom: string, day: string) {

    // First go through to see if any times overlap with existing data

    let tempUniqueID: number = 0;

    for (let i = 0; i < classRoomsTimes.length; i++) {

        // We use the .001 in this if to account for floating point errors
        if (startTime - .001 <= classRoomsTimes[i][0] && endTime - .001 >= classRoomsTimes[i][0]) {

            // Check to see if we're working with a new course, for formatting purposes
            if (uniqueIndex != tempUniqueID)
                startNewCourseBlock = true;
            else startNewCourseBlock = false;
            tempUniqueID = uniqueIndex;

            let roomHeadingColumnNumber: number;
            if (assignedRoom == "")
                assignedRoom = "Unassigned";
            for (let j = 0; j < roomHeadings[0].length; j++) {
                if (assignedRoom == roomHeadings[0][j]) {
                    roomHeadingColumnNumber = j + 1; // We add one because it is index 0 and starts at column B
                    for (let j = 0; j < days.length; j++) {
                        if ((day.includes(days[j]) && day != "Th") ||
                            (days[j] == "Th" && day == "Th")) {
                            PopulateDayAtTime(days[j], roomHeadingColumnNumber + j, i + 2, courseMnemonic, courseNumber, enrollment, uniqueIndex); // i + 1 because zero index and it starts at row 2
                        }
                    }
                }
            }
        }
    }
}



function PopulateDayAtTime(day: string, columnNumber: number, rowNumber: number, courseMnemonic: string, courseNumber: string, enrollment: number, uniqueIndex: number) {
    let cell = classRooms.getCell(rowNumber, columnNumber);

    var overlapFound: boolean = false;

    // Check for overlap
    for (let j = 0; j < cellHasValue.length; j++) {
        if (cellHasValue[j][0] == rowNumber && cellHasValue[j][1] == columnNumber) {
            overlapFound = true;
            break;
        }
        else {
            overlapFound = false;
        }
    }

    if (overlapFound) {
        let newCellValue = LookupStoredCellValue(rowNumber, columnNumber, courseMnemonic + " " + courseNumber + ": " + enrollment + " [" + uniqueIndex + "]");
        cell.setValue(newCellValue);
        cell.getFormat().getFill().setColor("red");
    }

    else {
        cell.setValue(courseMnemonic + " " + courseNumber + ": " + enrollment + " [" + uniqueIndex + "]");
        cell.getFormat().getFill().setColor(colors[colorIndex]);
    }

    cell.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.dashDot);

    if (startNewCourseBlock) {
        cell.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.dashDot);
    }



    cellHasValue.push([rowNumber, columnNumber]);
    cellValue.push(courseMnemonic + " " + courseNumber + ": " + enrollment + " [" + uniqueIndex + "]");

}

function LookupStoredCellValue(x: number, y: number, currentString: string) {
    for (let i = 0; i < cellValue.length; i++) {
        if (cellHasValue[i][0] == x && cellHasValue[i][1] == y) {
            cellValue[i] = cellValue[i] + " / " + currentString;
            return cellValue[i];
        }
    }
}

