function include(name){return HtmlService.createTemplateFromFile(name).getRawContent()}
function doGet(e){if(e&&e.parameter&&String(e.parameter.page||'')==='PublicKurs')return HtmlService.createTemplateFromFile('PublicKurs').evaluate().setTitle('Public Kurs').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);return MC_Router.render(e)}
function onOpen(){SpreadsheetApp.getUi().createMenu('MC-Almara').addItem('Setup System','setupSystem').addItem('Create Initial Admin','createInitialAdmin').addToUi()}
function setupSystem(){return MC_Database.install()}
function createInitialAdmin(){return MC_Auth.createInitialAdmin()}
function api(action,payload){return MC_Router.api(action,payload||{})}
function mcApi(action,payload){return MC_Router.api(action,payload||{})}
