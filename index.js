const argv = require('minimist')(process.argv.slice(2));
const rp = require('request-promise-native');
const moment = require('moment');
const XLSX = require('xlsx');

const config = require('./config.js');
const pivotalTrackerRESTToken = config.getPivotalTrackerRESTToken();
const pivotalTrackerProjectId = config.getPivotalTrackerProjectId();

function getReportURL(since, until) {
  var togglSummaryReportUrl = "https://toggl.com/reports/api/v2/summary?user_agent=node-toggl-export&workspace_id=" + config.getTogglWorkspaceId();

  var startDate = moment().format("YYYY-MM-DD");
  var endDate = moment().date(moment().date() + 1).format("YYYY-MM-DD");

  if (since) {
    startDate = moment(since, "YYYY-MM-DD").format("YYYY-MM-DD");
  }

  if (until) {
    endDate = moment(until, "YYYY-MM-DD").format("YYYY-MM-DD");
  }

  togglSummaryReportUrl += "&since=" + startDate;
  togglSummaryReportUrl += "&until=" + endDate;

  return togglSummaryReportUrl;
}

async function main() {

  try {
    var requestOptions = {
      "url": getReportURL(argv.since, argv.until),
      "auth": {
        "user": config.getTogglRESTToken(),
        "password": "api_token"
      },
      "headers": {
        "Content-Type": "application/json"
      },
      "json": true
    };

    console.log("Getting report from ", startDate, " to ", endDate);
    const summaryReport = await rp(requestOptions);
    console.log("Received good response from Toggl, processing data.");

    var taskPromises = [];

    summaryReport.data.forEach(function(project) {
      project.items.forEach(function(togglTask) {
        taskPromises.push(new Promise(function(accept, reject) {
          var task = {};
          var duration = moment.duration(togglTask.time);
          var hours = duration.asHours();
          task.duration = hours;
          task.type = project.title.project;
          task.name = togglTask.title.time_entry;
          task.owner = config.getYourName();

          try {
            if (task.name.indexOf("#") == 0) {
              task.id = task.name.split(" ")[0].replace("#", "");
              task.url = "https://www.pivotaltracker.com/story/show/" + task.id;
              var pivotalRequestPromiseOptions = {
                "uri": "https://www.pivotaltracker.com/services/v5/projects/" + pivotalTrackerProjectId + "/stories/" + task.id,
                "headers": {
                  "Content-Type": "application/json",
                  "X-TrackerToken": pivotalTrackerRESTToken
                },
                "json": true
              };
              rp(pivotalRequestPromiseOptions).then(function(pivotalTaskResponse) {
                task.estimate = pivotalTaskResponse.estimate;
                accept(task);
              }).catch(function(response) {
                console.error("Error while getting task from pivotal tracker: ");
                console.error(task.name);
                if (response && response.error) {
                  console.error(response.error.error);
                }
                accept(task);
              });
              return;
            }

            accept(task);

          } catch (e) {
            console.error(e);
            console.error("Error while processing " + task.name);
          }

        }));
      });

      const tasks = await Promise.all(taskPromises);
      console.log("All tasks processed succesfully.");
      saveTasksInExcelFile(tasks);

    });
  } catch (error) {
    console.error(error);
  }

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
    var fields = ["duration", "type", "owner", "estimate", "url", "name"];
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