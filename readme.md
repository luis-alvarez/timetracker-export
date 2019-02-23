Toggl/Pivotal Export
--------------------

This is just a quick script I made for exporting time tracked in Toggl.com/tracker of Pivotal Tasks.

To use it, create a config.json file that includes the restTokens and config for the pivotal tracker project and Toggl workspace.

EG config.json file

    {
      "pivotalTrackerRESTToken": "asdfasdfasdf",
      "pivotalTrackerProjectId": "1111111",
      "togglWorkspaceId": "11111",
      "togglRESTToken": "123123asdfasdf123123",
      "yourName": "Carlos"
    }

Run:
----
    node index.js //will export current date tasks to toggl-export.xslx

Available Params
----------------

	--since="YYYY-MM-DD"

	Allows you to specify an start date from which report should be generated. If not specified, today's date is used instead.

	--until="YYYY-MM-DD"

	Allows you to specify end date for the report date range. If not specified, tomorrow's date is used.

    If you ommit both params, you will get a today's time tracked report.


## Todo ##

 - Convert this to be a CLI
 - Add a weekly param to get a last week's report