function include(name){return HtmlService.createTemplateFromFile(name).getRawContent()}
function doGet(e){return MC_Router.render(e)}
function onOpen(){SpreadsheetApp.getUi().createMenu('MC-Almara').addItem('Setup System','setupSystem').addItem('Create Initial Admin','createInitialAdmin').addToUi()}
function setupSystem(){return MC_Database.install()}
function createInitialAdmin(){return MC_Auth.createInitialAdmin()}
function api(action,payload){return MC_Router.api(action,payload||{})}