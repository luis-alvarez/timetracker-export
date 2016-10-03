var config = require("./config.json");

function validateField(field) {
  if (!config[field]) {
    throw new Error("Check config.json, you need to specify a " + field + " property");
  }
}

module.exports = {
  "getPivotalTrackerRESTToken": function() {
    validateField("pivotalTrackerRESTToken");
    return config.pivotalTrackerRESTToken;
  },
  "getPivotalTrackerProjectId": function() {
    validateField("pivotalTrackerProjectId");
    return config.pivotalTrackerProjectId;
  },
  "getTogglWorkspaceId": function() {
    validateField("togglWorkspaceId");
    return config.togglWorkspaceId;
  },
  "getTogglRESTToken": function() {
    validateField("togglRESTToken");
    return config.togglRESTToken;
  },
  "getYourName": function() {
    validateField("yourName");
    return config.yourName;
  }
};