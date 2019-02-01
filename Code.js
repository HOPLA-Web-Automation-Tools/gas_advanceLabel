var scriptName = "Advanced Labeling";
var userProperties = PropertiesService.getUserProperties();

var labels = userProperties.getProperty("labels") || "";
var last_run = userProperties.getProperty("last_run") || "";

var user_email = Session.getEffectiveUser().getEmail();

global.doGet = doGet;
global.default_action = default_action;
global.deleteAllTriggers = deleteAllTriggers;
global.test = test;
global.main_trigger = main_trigger;
global.do_it_now = do_it_now;
global.main_doLabel = main_doLabel;
global.regex_subscription = regex_subscription;
global.markLabel = markLabel;
global.archive = archive;
global.threadHasLabel = threadHasLabel;
global.isMe = isMe;
global.getEmailAddresses = getEmailAddresses;
global.getLabel = getLabel;

function default_action() {
  deleteAllTriggers();

  var content = "<p>" + scriptName + " has been installed on your email " + user_email + ". "
    + '<p>It will:</p>'
    + '<ul style="list-style-type:disc">'
    + '<li>Label incoming emails with your specified labels when it passes the filters you specify.</li>'
    + '</ul>'
    + '<p>You can change these settings by clicking the HOPLA Tools extension icon</p>';

  return HtmlService.createHtmlOutput(content);
}

function doGet(e) {
  if (e.parameter.setup) { // SETUP
    return default_action();
  } else if (e.parameter.test) {
    var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    return HtmlService.createHtmlOutput("Status: " + authInfo.getAuthorizationStatus());
  } else if (e.parameter.savesettings && e.parameter.labels) { // SAVE SETTINGS
    userProperties.setProperty('labels', e.parameter.labels);
    var oLabels = JSON.parse(e.parameter.labels);


    deleteAllTriggers();

    var main_enable = false;

    for (var key in oLabels) {
      var label = oLabels[key];
      if (label.status === 'enabled') {
        main_enable = true;
      }
    }

    if (main_enable) {
      ScriptApp.newTrigger("main_trigger").timeBased().everyMinutes(5).create();
    }

    return ContentService.createTextOutput("settings has been saved.");
  } else if (e.parameter.run) { // DO IT NOW
    try {
      var labeled = do_it_now();
      return ContentService.createTextOutput(labeled + " threads has been labeled.");
    } catch (e) {
      return ContentService.createTextOutput(e);
    }
  } else if (e.parameter.get_triggers) { // ENABLE
    var triggers = ScriptApp.getProjectTriggers();
    return ContentService.createTextOutput(triggers.length + " trigger(s).");
  } else if (e.parameter.subscriptions_getVariables) { // GET VARIABLES
    var labels = userProperties.getProperty("labels") || "";
    return ContentService.createTextOutput(labels);
  } else { // NO PARAMETERS
    return default_action();
  }
}


function deleteAllTriggers() {
  // DELETE ALL TRIGGERS
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  // DELETE ALL TRIGGERS***
}

function test() {
  Logger.log(labels);
}

function main_trigger() {
  if (!last_run) {
    last_run = {};
  } else {
    last_run = JSON.parse(last_run);
  }
  var d = new Date();


  var labels = userProperties.getProperty("labels") || "";
  if (labels) {
    labels = JSON.parse(labels);

    var labels_to_process = [];
    for (var key in labels) {
      var label = labels[key];
      if (!last_run[key] && label.status === 'enabled') {
        Logger.log("no last run and is enabled");
        labels_to_process.push(label.labelname);
        last_run[key] = d.getTime();
      } else if (shouldRun(last_run[key], label.frequency) && label.status === 'enabled') {
        Logger.log("has lastrun, shouldrun and is enabled");
        labels_to_process.push(label.labelname);
        last_run[key] = d.getTime();
      } else {
        Logger.log(key + " is not yet scheduled.");
      }
    }
    if (labels_to_process.length) main_doLabel(labels_to_process);
  }

  userProperties.setProperty('last_run', JSON.stringify(last_run));

  function shouldRun(pLastRun, pFrequency) {
    return ((Number(pLastRun) + (label.frequency * 60 * 1000)) <= d.getTime());
  }
}

function do_it_now() {
  var labels = userProperties.getProperty("labels") || "";
  if (labels) {
    labels = JSON.parse(labels);

    var labels_to_process = [];
    for (var key in labels) {
      var label = labels[key];
      if (label.status === 'enabled') {
        labels_to_process.push(label.labelname);
      }
    }
    if (labels_to_process.length) return main_doLabel(labels_to_process);

    throw new Error("No advance label is enabled.");
  }
  throw new Error("No advance labels saved.");
}

function main_doLabel(aLabelsToProcess) {
  if (!labels) labels = userProperties.getProperty('labels');
  if (!labels) return;
  labels = JSON.parse(labels);

  var d = new Date();
  var minutesAgo = 120; // 2 HOURS AGO
  var n = ((d.getTime() / 1000) - (minutesAgo * 60));
  n = n.toFixed();
  Logger.log("n:" + n);

  var filters = [
    'in:inbox',
    'after:' + n
  ];

  try {
    var threads = GmailApp.search(filters.join(' '));
  } catch (e) {
    deleteAllTriggers();
    return;
  }

  var threadMessages = GmailApp.getMessagesForThreads(threads);

  Logger.log("# of threads found for the past " + minutesAgo + " minutes: " + threadMessages.length);
  var count_pass = 0,
    count_fail = 0;

  var ThreadsToLabel = {};

  for (var i = 0; i < threadMessages.length; i++) {
    var lastMessage = threadMessages[i][threadMessages[i].length - 1],
      lastFrom = lastMessage.getFrom(),
      body = lastMessage.getRawContent(),
      subject = lastMessage.getSubject(),
      thread = lastMessage.getThread();

    for (var key in labels) {
      var label = labels[key];
      if (aLabelsToProcess.indexOf(label.labelname) > -1) { // IF THIS LABEL IS TO PROCESS.
        if (filter_check(subject, body, label.filters, label.case_sensitive, label.use_regex)) {
          Logger.log("LABELED: " + subject);
          add_to_threadstolabel(thread, label.labelname);
          count_pass += 1;
        } else {
          Logger.log("NOT LABELED: " + subject);
          count_fail += 1;
        }
      }
    }
  }

  var msg = "Labeled Threads=" + count_pass + " NOT Labeled threads:" + count_fail;
  Logger.log(msg);


  markLabel(ThreadsToLabel);
  archive(ThreadsToLabel);

  return count_pass;


  function filter_check(pSubject, pBody, aFilters, bCaseSensitive, bRegex) {
    for (var index in aFilters) {
      var filter = aFilters[index];
      if (!bCaseSensitive) { // IF NOT CASE SENSITIVE
        filter = filter.toLowerCase();
        pSubject = pSubject.toLowerCase();
        pBody = pBody.toLowerCase();
      }

      if (!bRegex) { // NOT REGEX
        if (pSubject.indexOf(filter) > -1) {
          Logger.log("Subject of '" + pSubject + "' contains '" + filter + "'");
          return true;
        }
        if (pBody.indexOf(filter) > -1) {
          Logger.log("Body of '" + pSubject + "' contains '" + filter + "'");
          return true;
        }
      } else { // REGEX
        var rgx = new RegExp(filter, 'g');
        var match = pSubject.match(rgx);
        if (match) {
          Logger.log("Subject of '" + pSubject + "' matches regex '" + filter + "'");
          return true;
        }
        match = pBody.match(rgx);
        if (match) {
          Logger.log("Body of '" + pSubject + "' matches regex '" + filter + "'");
          return true;
        }
      }
    }
    return false;
  }

  function add_to_threadstolabel(thread, labelname) {
    if (!ThreadsToLabel[labelname]) {
      ThreadsToLabel[labelname] = [thread];
    } else {
      ThreadsToLabel[labelname].push(thread);
    }
  }
}


function regex_subscription(pBody, pSubject) {
  pBody = pBody.replace(/3D"/g, '"');
  pBody = pBody.replace(/=\s/g, '');
  var urls = pBody.match(/^list\-unsubscribe:(.|\r\n\s)+<(https?:\/\/[^>]+)>/im);
  if (urls) {
    Logger.log("Subject: " + pSubject + " Unsubscribe link: " + urls[2]);
    return 1;
  }

  urls = pBody.match(/^list\-unsubscribe:(.|\r\n\s)+<mailto:([^>]+)>/im);
  if (urls) {
    Logger.log("Subject: " + pSubject + " Unsubscribe email: " + urls[2]);
    return 1;
  }

  // Regex to find all hyperlinks
  var hrefs = new RegExp(/<a[^>]*href=["'](https?:\/\/[^"']+)["'][^>]*>(.*?)<\/a>/gi);

  // Iterate through all hyperlinks inside the message
  while (urls = hrefs.exec(pBody)) {
    // Does the anchor text or hyperlink contain words like unusbcribe or optout
    if (urls[1].match(/unsubscribe|optout|opt\-out|remove/i) || urls[2].match(/unsubscribe|optout|opt\-out|remove/i)) {
      Logger.log("Subject: " + pSubject + " HTML unsubscribe link: " + urls[1]);
      return 1;
    }
  }

  return 0;
}

function markLabel(pThreadsToLabel) {
  var ADD_LABEL_TO_THREAD_LIMIT = 100;

  for (var labelname in pThreadsToLabel) {
    var threads = pThreadsToLabel[labelname];
    Logger.log("labelname: " + labelname);
    var oLabel = getLabel(labelname);
    Logger.log("oLabel: " + oLabel);
    if (threads.length > ADD_LABEL_TO_THREAD_LIMIT) {
      for (var i = 0; i < Math.ceil(threads.length / ADD_LABEL_TO_THREAD_LIMIT); i++) {
        oLabel.addToThreads(threads.slice(100 * i, 100 * (i + 1)));
      }
    } else {
      oLabel.addToThreads(threads);
    }
  }
}

function archive(pThreadsToLabel) {
  for (var labelname in pThreadsToLabel) {
    var threads = pThreadsToLabel[labelname];
    for (var i = 0; i < threads.length; i++) {
      threads[i].moveToArchive();
    }
  }
}

function threadHasLabel(thread, labelName) {
  var labels = thread.getLabels();
  for (var i = 0; i < labels.length; i++) {
    var label = labels[i];
    if (label.getName() === labelName) {
      return true;
    }
  }
  return false;
}

function isMe(fromAddress) {
  var addresses = getEmailAddresses();
  for (var i = 0; i < addresses.length; i++) {
    var address = addresses[i],
      r = RegExp(address, 'i');

    if (r.test(fromAddress)) {
      return true;
    }
  }

  return false;
}

function getEmailAddresses() {
  // Cache email addresses to cut down on API calls.
  if (!this.emails) {
    var me = Session.getActiveUser().getEmail(),
      emails = GmailApp.getAliases();

    emails.push(me);
    this.emails = emails;
  }
  return this.emails;
}

function getLabel(labelName) {
  // Cache the labels.
  this.labelsObjects = this.labelsObjects || {};
  label = this.labelsObjects[labelName];

  if (!label) {
    // Logger.log('Could not find cached label "' + labelName + '". Fetching.', this.labels);

    var label = GmailApp.getUserLabelByName(labelName);

    if (label) {
      // Logger.log('Label exists.');
    } else {
      // Logger.log('Label does not exist. Creating it.');
      label = GmailApp.createLabel(labelName);
    }
    this.labelsObjects[labelName] = label;
  }
  return label;
}