var argv = require('minimist')(process.argv.slice(2));
var excelbuilder = require('msexcel-builder');
var rp = require('request-promise');
var config = require('./config.js');
var moment = require('moment');
var Promise = require('promise');

var pivotalTrackerRESTToken = config.getPivotalTrackerRESTToken();
var pivotalTrackerProjectId = config.getPivotalTrackerProjectId();

var togglSummaryReportUrl = "https://toggl.com/reports/api/v2/summary?user_agent=node-toggl-export&workspace_id=" + config.getTogglWorkspaceId();

if (argv.since) {
  togglSummaryReportUrl += "&since=" + argv.since;
}

if (argv.until) {
  togglSummaryReportUrl += "&until=" + argv.until;
}

var requestOptions = {
  "url": togglSummaryReportUrl,
  "auth": {
    "user": config.getTogglRESTToken(),
    "password": "api_token"
  },
  "headers": {
    "Content-Type": "application/json"
  },
  "json": true
};

rp(requestOptions).then(function (summaryReport) {
  console.log("Received good response from Toggl, processing data.");
  var taskPromises = [];

  summaryReport.data.forEach(function(project) {
    project.items.forEach(function(togglTask) {
      taskPromises.push(new Promise(function(accept, reject) {
        var task = {};
        var duration = moment.duration(togglTask.time);
        var hours = Math.floor(duration.asHours());
        var minutes = duration.minutes();
        task.duration = hours + ":" + minutes;
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

  });

  Promise.all(taskPromises).then(function(tasks) {
    console.log("All tasks processed succesfully.");
    saveTasksInExcelFile(tasks);
  }).catch(function(error) {
    console.log("Error while getting/processing all the tasks");
    console.error(error);
  });

}).catch(function (err) {
  console.error(err);
});

function saveTasksInExcelFile(tasks) {
  var fields = ["duration", "type", "owner", "estimate", "url", "name"];
  // Create a new workbook file in current working-path
  var workbook = excelbuilder.createWorkbook('./', 'tasks.xlsx')

  // Create a new worksheet with 10 columns and 12 rows
  var sheet1 = workbook.createSheet('sheet1', fields.length, tasks.length);

    for (var i = 0, j = tasks.length; i < j; i++) {
      for (var k = 0, l = fields.length; k < l; k++) {
        try {
          var field = fields[k];
          sheet1.set(k + 1, i + 1, tasks[i][field]);
        } catch (error) {
          console.error(error);
        }
      }
    }

  workbook.save(function (ok) {
    if (!ok)
      workbook.cancel();
    else
      console.log('congratulations, your taks have been exported succesfully');
  });
}