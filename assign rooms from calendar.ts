function main(workbook: ExcelScript.Workbook) {
    // Your code here

  var courses = workbook.getWorksheet("Courses");

  var calendar = workbook.getActiveWorksheet();

  if (calendar.getName() == "Courses") {
    throw "Only run this on calendar pages";

  }
  let assignedRooms: string[][] = calendar.getRange("B1:M1").getValues() as string[][];
  let classes: string[][] = calendar.getRange("B2:M54").getValues() as string[][];
  let uniqueIndexes: string[][] = courses.getRange("O4:O200").getValues() as string[][];
  // simplify this array
  let simpleUniqueIndexes: string[] = [];
  for (let i = 0; i < uniqueIndexes.length; i++)
  {
    simpleUniqueIndexes[i] = String(uniqueIndexes[i][0]);
  }
  
  for(let i = 0; i < classes.length; i++)
  {
    for (let j = 0; j < classes[i].length; j++)
    {
      let cellValue = classes[i][j];
      let indexOfFirstBracket = cellValue.indexOf("[");
      let indexOfLastBracket = cellValue.indexOf("]");
      let uniqueIndex = cellValue.slice(indexOfFirstBracket + 1, indexOfLastBracket);
      let uniqueIndexString = String(uniqueIndex);
      

      let classIndex = simpleUniqueIndexes.indexOf(uniqueIndexString);
      courses.getCell(classIndex + 3, 12).setValue(assignedRooms[0][j]);
    }
    
  }

}