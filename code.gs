function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Classroom Name Extractor')
    .addItem('Get all the names!', 'getAllNames')
    .addToUi();
}

// Helper function to get the list of students in the course
function getStudents(courseId) {
  const optionalArgs = {
    pageSize: 100
  };
  let students = [];
  let response = Classroom.Courses.Students.list(courseId, optionalArgs);
  let studentsList = response.students;
  if (studentsList && studentsList.length > 0) {
    students = studentsList;
  }
  return students;
}

function getAllNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Clear existing content in the sheet
  sheet.clear();

  // Fetch the list of courses from Google Classroom
  const courses = Classroom.Courses.list();

  // Check if courses are available
  if (courses.courses && courses.courses.length > 0) {
    const data = [];
    
    // Filter only active courses
    const activeCourses = courses.courses.filter(c => c.courseState === "ACTIVE");
    
    // Fetch students for each active course and prepare the data array
    for (let i = 0; i < activeCourses.length; i++) {
      const course = activeCourses[i];
      const courseId = course.id;
      const courseName = course.name;
      
      // Get the list of students for the course
      const students = getStudents(courseId);
      
      // Prepare the header row with the course name
      const columnData = [courseName];

      // Add student names to the column
      for (let j = 0; j < students.length; j++) {
        columnData.push(students[j].profile.name.givenName);
      }

      // Add the column data to the data array
      data.push(columnData);
    }
    
    if (data.length > 0) {
      // Transpose data to fit into columns
      const transposedData = transposeArray(data);
      
      // Set the values in the sheet starting from the first row and column
      sheet.getRange(1, 1, transposedData.length, transposedData[0].length).setValues(transposedData);
    } else {
      sheet.getRange(1, 1).setValue("No active courses found");
    }
  } else {
    // If no courses are found, add a message row
    sheet.getRange(1, 1).setValue("No courses found");
  }
}

// Helper function to transpose a 2D array considering the longest column
function transposeArray(array) {
  const maxLength = Math.max(...array.map(col => col.length));
  const newArray = [];
  
  for (let i = 0; i < maxLength; i++) {
    newArray[i] = [];
    for (let j = 0; j < array.length; j++) {
      newArray[i][j] = array[j][i] || "";  // Fill with empty string if there is no data
    }
  }
  return newArray;
}
