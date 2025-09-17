/**
 * Google Classroom Name Extractor
 *
 * File-level overview:
 * This script integrates Google Classroom with Google Sheets to extract student
 * given names from all active courses and write them into the active sheet,
 * one course per column with the course name as the header.
 *
 * Required setup:
 * - Enable the Classroom Advanced Service in the Apps Script project (Resources > Advanced Google services).
 * - Enable the Google Classroom API in the Google Cloud Console for the project.
 * - Required OAuth scopes (at minimum):
 *     https://www.googleapis.com/auth/classroom.courses.readonly
 *     https://www.googleapis.com/auth/classroom.rosters.readonly
 *   (The script also uses SpreadsheetApp; Apps Script will add the necessary spreadsheet scopes.)
 *
 * Usage:
 * - Install the script into a Google Sheets container-bound Apps Script project.
 * - Reload the spreadsheet to trigger onOpen(), which adds the "Classroom Name Extractor" menu.
 * - Use the menu item "Get all the names!" or run getAllNames() directly from the Apps Script editor.
 *
 * Side effects:
 * - getAllNames() will clear the entire active sheet (sheet.clear()) before writing output.
 * - The script writes a transposed 2D array so that each active course becomes a column:
 *     - The top cell of each column is the course name.
 *     - Subsequent cells in the column are students' given names (profile.name.givenName).
 * - If no courses or no active courses are found, the script writes a single message cell:
 *     - "No courses found" or "No active courses found" respectively.
 *
 * Functions:
 *
 * onOpen()
 * - Adds a custom menu ("Classroom Name Extractor") to the Google Sheets UI with one item:
 *     "Get all the names!" which calls getAllNames().
 * - No parameters or return value.
 *
 * getStudents(courseId)
 * - Fetches the list of students for the given courseId using Classroom.Courses.Students.list.
 * - Handles pagination by iterating with nextPageToken until all pages are collected.
 * - Uses a pageSize of 100 per request.
 * - Parameters:
 *     - courseId (string): the Classroom course ID to fetch students for.
 * - Returns:
 *     - Array of student objects (may be empty if there are no students).
 * - Notes:
 *     - Student objects are returned as provided by the Classroom API (e.g., student.profile.name).
 *     - API errors (authorization, rate limits, etc.) will surface as exceptions.
 *
 * getAllNames()
 * - Primary workflow:
 *     1. Clears the active sheet.
 *     2. Fetches all courses via Classroom.Courses.list().
 *     3. Filters courses to courseState === "ACTIVE".
 *     4. For each active course, retrieves all students using getStudents(courseId).
 *     5. Builds a column array for each course: [courseName, studentGivenName1, ...].
 *     6. Transposes the columns into rows (so each course becomes a sheet column).
 *     7. Writes the transposed array to the sheet starting at A1.
 * - Behavior on empty results:
 *     - If no courses are returned => writes "No courses found" to A1.
 *     - If courses exist but none are ACTIVE => writes "No active courses found" to A1.
 * - Side effects:
 *     - Clears and overwrites the active sheet.
 * - No return value.
 *
 * transposeArray(array)
 * - Utility to transpose an array of columns into rows.
 * - Accepts columns of varying lengths (ragged arrays) and pads missing entries with empty strings.
 * - Parameters:
 *     - array (Array<Array<any>>): an array where each element is a column array.
 * - Returns:
 *     - Array<Array<any>>: the transposed 2D array ready for setValues().
 *
 * Error handling and limits:
 * - The script does not include explicit retry logic for transient errors or quota handling.
 * - Classroom API quotas and rate limits apply. Consider exponential backoff for production use.
 * - If profile.name.givenName is undefined for a student, the cell will be written as an empty string.
 *
 * Security and privacy:
 * - The script reads student names from Classroom. Ensure you have appropriate permissions
 *   and follow your organization's privacy policies before exporting roster data to Sheets.
 */
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
  let pageToken = null;

  do {
    if (pageToken) {
      optionalArgs.pageToken = pageToken;
    } else {
      delete optionalArgs.pageToken;
    }

    const response = Classroom.Courses.Students.list(courseId, optionalArgs);
    const studentsList = response && response.students ? response.students : [];

    // Append students from this page
    for (let i = 0; i < studentsList.length; i++) {
      students.push(studentsList[i]);
    }

    // Prepare for next iteration
    pageToken = response && response.nextPageToken ? response.nextPageToken : null;
  } while (pageToken);

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
