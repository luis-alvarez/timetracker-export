const argv = require('minimist')(process.argv.slice(2));
const rp = require('request-promise-native');
const moment = require('moment');
const XLSX = require('xlsx');

const config = require('./config.js');
const pivotalTrackerRESTToken = config.getPivotalTrackerRESTToken();
const pivotalTrackerProjectId = config.getPivotalTrackerProjectId();

const { uniq: removeDuplicates, keyBy, groupBy, keys: getKeys } = require("lodash");

const TOGGL_TIME_ENTRIES_URL = "https://www.toggl.com/api/v8/time_entries?user_agent=node-toggl-export&workspace_id=";
const TOGGL_PROJECT_URL = "https://www.toggl.com/api/v8/projects/";
const OWNER_NAME = "Luis Alvarez";

function getTimeEntriesURL(since, until) {
  var togglTimeEntriesUrl = TOGGL_TIME_ENTRIES_URL + config.getTogglWorkspaceId();

  var startDate = moment().toISOString();
  var endDate = moment().date(moment().date() + 1).toISOString();

  if (since) {
    startDate = moment(since, "YYYY-MM-DD").toISOString();
  }

  if (until) {
    endDate = moment(until, "YYYY-MM-DD").toISOString();
  }

  togglTimeEntriesUrl += "&start_date=" + startDate;
  togglTimeEntriesUrl += "&end_date=" + endDate;

  return togglTimeEntriesUrl;
}

function getTimeEntries() {
  const requestOptions = {
    "url": getTimeEntriesURL(argv.since, argv.until),
    "auth": {
      "user": config.getTogglRESTToken(),
      "password": "api_token"
    },
    "headers": {
      "Content-Type": "application/json"
    },
    "json": true
  };
  return rp(requestOptions);
}

function getProject (projectId) {
  const requestOptions = {
    "url": TOGGL_PROJECT_URL + projectId,
    "auth": {
      "user": config.getTogglRESTToken(),
      "password": "api_token"
    },
    "headers": {
      "Content-Type": "application/json"
    },
    "json": true
  };
  return rp(requestOptions);
}

async function getProjectsFromIds(projectIds) {
  const projects = (await Promise.all(projectIds.filter(Boolean).map(pid => getProject(pid)))).map(project => project.data);
  return keyBy(projects, "id");
}

const getProjectIdsFromTimeEntries = (timeEntries) => removeDuplicates(timeEntries.map(te => te.pid));

const getPivotalTrackerData = async (taskName) => {
  try {
    if (taskName.indexOf("#") == 0) {
      const pivotalId = taskName.split(" ")[0].replace("#", "");
      const pivotalTaskURL = "https://www.pivotaltracker.com/story/show/" + pivotalId;

      const pivotalTaskRPOptions = {
        "uri": "https://www.pivotaltracker.com/services/v5/projects/" + pivotalTrackerProjectId + "/stories/" + pivotalId,
        "headers": {
          "Content-Type": "application/json",
          "X-TrackerToken": pivotalTrackerRESTToken
        },
        "json": true
      };

      const pivotalTask = await rp(pivotalTaskRPOptions);

      return {
        "estimate": pivotalTask.estimate,
        "url": pivotalTaskURL
      };
    }
    return {};
  } catch (e) {
    console.error(e);
    console.error("Error while getting information from pivotal tracker for task: " + taskName);
  }
  return {};
};

const getTaskData = (groupedTimeEntries, projects) => {
  return async (taskId) => {
    const taskTimeEntries = groupedTimeEntries[taskId];
    const task = taskTimeEntries[0];
    const taskTotalDuration = taskTimeEntries.reduce((totalDuration, { duration }) => totalDuration + duration, 0);
    const daysWorked = removeDuplicates(taskTimeEntries.map(te => moment(te.start).format("dddd"))).join(", ");
    const pivotalData = await getPivotalTrackerData(task.description);

    return  {
      "name": task.description,
      daysWorked,
      ...pivotalData,
      "duration": moment.duration(taskTotalDuration, 'seconds').asHours(),
      "type": projects[task.pid].name,
      "owner": OWNER_NAME,
      "notes": ""
    }
  };
}

async function main() {

  const timeEntries = await getTimeEntries();
  const projecIds = getProjectIdsFromTimeEntries(timeEntries);
  const projects = await getProjectsFromIds(projecIds);
  console.log("Received good response from Toggl, processing data.");

  const groupedTimeEntries = groupBy(timeEntries, "description");
  const performedTasksIds = getKeys(groupedTimeEntries);
  const performedTasks = await Promise.all(performedTasksIds.map(getTaskData(groupedTimeEntries, projects)));

  console.log("All tasks processed succesfully.");

  saveTasksInExcelFile(performedTasks);

}


function Workbook() {
	if(!(this instanceof Workbook)) return new Workbook();
	this.SheetNames = [];
	this.Sheets = {};
}

function datenum(v, date1904) {
	if(date1904) v+=1462;
	var epoch = Date.parse(v);
	return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function sheet_from_array_of_arrays(data, fields, opts) {
	var ws = {};
	var range = {s: {c:fields.length, r:data.length}, e: {c:0, r:0 }};
	for(var R = 0; R != data.length; ++R) {
		for(var C = 0; C != data[R].length; ++C) {
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			var cell = {v: data[R][C] };
			if(cell.v == null) continue;
			var cell_ref = XLSX.utils.encode_cell({c:C,r:R});

			if(typeof cell.v === 'number') {
        cell.t = 'n';
        if (C == 0) {
          cell.z = "#,00"
        }
        XLSX.utils.format_cell(cell);
      }
			else if(typeof cell.v === 'boolean') cell.t = 'b';
			else if(cell.v instanceof Date) {
				cell.t = 'n'; cell.z = XLSX.SSF._table[14];
				cell.v = datenum(cell.v);
			}
			else cell.t = 's';

			ws[cell_ref] = cell;
		}
	}
	if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
	return ws;
}

function saveTasksInExcelFile(tasks) {
  try {
    // Create the dataSet from the tasks and fields
    var fields = ["duration", "type", "owner", "estimate", "url", "name", "notes", "daysWorked"];
    var tasksRows = [];
    tasks.forEach(function(task) {
      var taskColumns = [];
      fields.forEach(function(field) {
        taskColumns.push(task[field]);
      });
      tasksRows.push(taskColumns);
    });
    var ws_name = "Toggl Export";

    var wb = new Workbook(), ws = sheet_from_array_of_arrays(tasksRows, fields);

    /* add worksheet to workbook */
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;

    /* write file */
    XLSX.writeFile(wb, 'toggl-export.xlsx');
  } catch (error) {
    console.error(error);
  }
}

main()
.then(() => {
  console.log("all done!");
})
.catch((error) => {
  console.error(error);
});