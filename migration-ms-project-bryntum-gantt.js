import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { dirname } from "path";
import xlsx from "xlsx";

const excelfilePath = "./launch-website.xlsx";
const projectName = "Launch website";

function convertToDate(excelDate) {
  // Convert Excel serial date to JavaScript Date
  // Excel dates are days since 1/1/1900
  const excelEpoch = new Date(1899, 11, 30);
  const jsDate = new Date(excelEpoch.getTime() + excelDate * 86400000);
  // Return date in YYYY-MM-DD format without time
  return jsDate.toISOString().split('T')[0];
}

function parseDuration(durationStr) {
  // Parse duration strings like "3 days", "2 days" and return number of days
  if (!durationStr || typeof durationStr !== 'string') return 0;
  const match = durationStr.match(/(\d+)\s*(day|days)/i);
  return match ? parseInt(match[1]) : 0;
}

function calculateDuration(startDate, endDate) {
  // Calculate duration in days from start to end date
  // Bryntum duration is the number of days to ADD to start date to get end date
  // So for Nov 11 to Nov 13, duration should be 2 (11 + 2 = 13)
  if (!startDate || !endDate) return 0;

  const start = new Date(startDate);
  const end = new Date(endDate);

  // Set both to start of day for accurate day counting
  start.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);

  // Calculate difference in milliseconds and convert to days
  const diffTime = end - start;
  const diffDays = diffTime / (1000 * 60 * 60 * 24);

  return diffDays;
}

function parseDependencyString(depString) {
  // Parse dependency strings like "1FS", "2SS", "3FF,4SF" etc.
  // Returns array of {outlineNumber, type}
  if (!depString) return [];

  const deps = depString.toString().split(',').map(d => d.trim()).filter(d => d);
  return deps.map(dep => {
    // Extract outline number and dependency type (e.g., "1FS" -> {outlineNumber: "1", type: "FS"})
    const match = dep.match(/^([\d.]+)([A-Z]{2})$/);
    if (match) {
      return {
        outlineNumber: match[1],
        type: match[2]
      };
    }
    return null;
  }).filter(d => d !== null);
}

function mapDependencyType(msProjectType) {
  // Map Microsoft Project dependency types to Bryntum
  const typeMap = {
    'FS': 2,
    'SS': 0,
    'FF': 3,
    'SF': 1 
  };
  return typeMap[msProjectType] || 2; // Default to FS if unknown
}

function createBryntumTasksRows(data) {
  let taskId = 0;
  const tree = [];
  const taskMap = {}; // Map outline numbers to task objects
  const rawDependencies = []; // Store raw dependencies to process later

  // Start from index 8 where actual task data begins (after header row)
  for (let i = 8; i < data.length; i++) {
    const row = data[i];

    // Skip empty rows
    if (!row["__EMPTY"]) continue;

    const outlineNumber = row[projectName];
    const name = row["__EMPTY"];
    const assignedTo = row["__EMPTY_1"] || undefined;
    const startDate = row["__EMPTY_2"] ? convertToDate(row["__EMPTY_2"]) : undefined;
    const endDate = row["__EMPTY_3"] ? convertToDate(row["__EMPTY_3"]) : undefined;
    const duration = startDate && endDate ? calculateDuration(startDate, endDate) : parseDuration(row["__EMPTY_4"]);
    const bucket = row["__EMPTY_5"];
    const percentComplete = row["__EMPTY_6"] ? row["__EMPTY_6"] * 100 : 0;
    const priority = row["__EMPTY_7"];
    const dependsOn = row["__EMPTY_8"]; // Dependencies column
    const milestone = row["__EMPTY_13"] === "Yes";
    const notes = row["__EMPTY_14"];

    // Determine if this is a parent task (outline number has no decimal)
    const isParent = !outlineNumber.toString().includes(".");

    const task = {
      id: taskId,
      name,
      ...(startDate && { startDate }),
      ...(duration !== undefined && { duration }),
      ...(assignedTo && { resourceAssignment: assignedTo }),
      percentDone: percentComplete,
      ...(isParent && { expanded: true }),
      ...(milestone && { duration: 0 }),
      ...(notes && { note: notes }),
      ...(bucket && { bucket }),
      ...(priority && { priority }),
      children: []
    };

    // Store raw dependencies for later processing
    if (dependsOn) {
      const deps = parseDependencyString(dependsOn);
      if (deps.length > 0) {
        rawDependencies.push({
          taskId: taskId,
          outlineNumber: outlineNumber,
          dependencies: deps
        });
      }
    }

    // Add to tree or to parent's children
    if (isParent) {
      tree.push(task);
    } else {
      // Extract parent outline number (e.g., "2.1" -> "2")
      const parentOutline = outlineNumber.toString().split(".")[0];
      const parent = taskMap[parentOutline];
      if (parent) {
        parent.children.push(task);
      }
    }

    // Store mapping for children to find their parent
    taskMap[outlineNumber] = task;

    taskId++;
  }

  return { tree, rawDependencies, taskMap };
}

const workbook = xlsx.readFile(excelfilePath);
const sheetName = workbook.SheetNames[0]; 
const worksheet = workbook.Sheets[sheetName];

const jsonData = xlsx.utils.sheet_to_json(worksheet);
const { tree: tasksRows, rawDependencies, taskMap } = createBryntumTasksRows(jsonData);

// Create dependencies array
const dependenciesRows = [];
let dependencyId = 0;

for (const rawDep of rawDependencies) {
  for (const dep of rawDep.dependencies) {
    const fromTask = taskMap[dep.outlineNumber];
    const toTask = taskMap[rawDep.outlineNumber];

    if (fromTask && toTask) {
      dependenciesRows.push({
        id: dependencyId++,
        fromTask: fromTask.id,
        toTask: toTask.id,
        type: mapDependencyType(dep.type)
      });
    }
  }
}

// Convert JSON data to the expected load response structure
const ganttLoadResponse = {
  success: true,
  tasks: {
    rows: tasksRows,
  },
  dependencies: {
    rows: dependenciesRows,
  },
};

let dataJson = JSON.stringify(ganttLoadResponse, null, 2); // Convert the data to JSON, indented with 2 spaces

// Define the path to the JSON file
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const filePath = path.join(__dirname, "data.json");

// Write the JSON string to a file
fs.writeFile(filePath, dataJson, (err) => {
  if (err) throw err;
  console.log("JSON data written to data.json");
});