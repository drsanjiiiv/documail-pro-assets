/**
 * ============================================================================
 * DOCUMAIL PRO COMPLETE MASTER CORE SCRIPT
 * ============================================================================
 * FILE: triggers.gs - Time-Driven Trigger Management
 * ============================================================================
 */

var TRIGGER_MAP_PREFIX = 'DOCUMAIL_TRIGGER_TEMPLATE_';
var TRIGGER_UID_PREFIX = 'DOCUMAIL_TRIGGER_UID_';

function GET_SCHEDULE_MINUTES(scheduleValue) {
  var map = { '5': 5, '15': 15, '60': 60, '360': 360, '720': 720, '1440': 1440 };
  return map[String(scheduleValue)] || null;
}

function MANAGE_TRIGGER_FOR_TEMPLATE(templateId) {
  try {
    var template = GET_TEMPLATE_BY_ID(templateId);
    if (!template) {
      Logger.log('MANAGE_TRIGGER_FOR_TEMPLATE: Template not found, removing trigger');
      REMOVE_TEMPLATE_TRIGGER(templateId);
      return;
    }

    var schedule = template.schedule || 'MANUAL';
    if (schedule === 'MANUAL') {
      Logger.log('MANAGE_TRIGGER_FOR_TEMPLATE: Schedule is MANUAL, removing trigger');
      REMOVE_TEMPLATE_TRIGGER(templateId);
      return;
    }

    var minutes = GET_SCHEDULE_MINUTES(schedule);
    if (minutes === null) {
      Logger.log('MANAGE_TRIGGER_FOR_TEMPLATE: Invalid schedule value: ' + schedule);
      REMOVE_TEMPLATE_TRIGGER(templateId);
      return;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = ss.getActiveSheet().getName();

    REMOVE_TEMPLATE_TRIGGER(templateId);

    var triggerBuilder = ScriptApp.newTrigger('RUN_SCHEDULED_TEMPLATES').timeBased();
    if (minutes < 60) {
      triggerBuilder.everyMinutes(minutes);
    } else {
      triggerBuilder.everyHours(minutes / 60);
    }
    var trigger = triggerBuilder.create();
    Logger.log('Trigger created for template: ' + template.name + ' (every ' + minutes + ' min)');

    var mapping = {
      templateId: templateId,
      sheetName: sheetName,
      spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
      triggerUid: trigger.getUniqueId(),
      template: template
    };

    PropertiesService.getUserProperties().setProperty(
      TRIGGER_MAP_PREFIX + templateId, JSON.stringify(mapping)
    );
    PropertiesService.getUserProperties().setProperty(
      TRIGGER_UID_PREFIX + trigger.getUniqueId(), templateId
    );

    console.log('Trigger stored for template: ' + template.name);
  } catch (e) {
    Logger.log('MANAGE_TRIGGER_FOR_TEMPLATE error: ' + e.message);
  }
}

function REMOVE_TEMPLATE_TRIGGER(templateId) {
  try {
    var mapKey = TRIGGER_MAP_PREFIX + templateId;
    var mappingJson = PropertiesService.getUserProperties().getProperty(mapKey);
    if (!mappingJson) return;

    var mapping = JSON.parse(mappingJson);

    PropertiesService.getUserProperties().deleteProperty(TRIGGER_UID_PREFIX + mapping.triggerUid);

    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getUniqueId() === mapping.triggerUid) {
        ScriptApp.deleteTrigger(triggers[i]);
        console.log('Trigger removed for template: ' + templateId);
        break;
      }
    }
  } catch (e) {
    console.log('REMOVE_TEMPLATE_TRIGGER error: ' + e.message);
  } finally {
    try {
      PropertiesService.getUserProperties().deleteProperty(TRIGGER_MAP_PREFIX + templateId);
    } catch (ex) {}
  }
}

function RUN_SCHEDULED_TEMPLATES(e) {
  try {
    if (!e || !e.triggerUid) return;

    var uidKey = TRIGGER_UID_PREFIX + e.triggerUid;
    var templateId = PropertiesService.getUserProperties().getProperty(uidKey);
    if (!templateId) return;

    var mapKey = TRIGGER_MAP_PREFIX + templateId;
    var mappingJson = PropertiesService.getUserProperties().getProperty(mapKey);
    if (!mappingJson) return;

    var mapping = JSON.parse(mappingJson);
    if (!mapping || !mapping.sheetName || !mapping.templateId || !mapping.spreadsheetId) return;

    var ss = SpreadsheetApp.openById(mapping.spreadsheetId);
    var sheet = ss.getSheetByName(mapping.sheetName);
    if (!sheet) return;

    RUN_TEMPLATE(mapping.templateId, sheet, mapping.template);

  } catch (err) {
    console.log('RUN_SCHEDULED_TEMPLATES error: ' + err.message);
  }
}

function TEST_CREATE_TRIGGER() {
  try {
    var trigger = ScriptApp.newTrigger('TEST_HANDLER').timeBased().everyMinutes(5).create();
    Logger.log("✅ Trigger created: " + trigger.getUniqueId());
    ScriptApp.deleteTrigger(trigger);
    Logger.log("✅ Trigger deleted - authorization OK");
  } catch (e) {
    Logger.log("❌ ERROR creating trigger: " + e.message);
  }
}

function TEST_HANDLER() {}

function REAUTHORIZE_SCRIPT() {
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log("Found " + triggers.length + " existing triggers. Authorization OK.");
}
