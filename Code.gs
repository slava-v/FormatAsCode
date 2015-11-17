//   This file is part of FormatAsCode add-on
//
//    FormatAsCode add-onis free software: you can redistribute it and/or modify
//    it under the terms of the GNU General Public License as published by
//    the Free Software Foundation, either version 3 of the License, or
//    (at your option) any later version.
//
//    FormatAsCode add-on is distributed in the hope that it will be useful,
//    but WITHOUT ANY WARRANTY; without even the implied warranty of
//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//    GNU General Public License for more details.
//
//    You should have received a copy of the GNU General Public License
//    along with FormatAsCode add-on. If not, see <http://www.gnu.org/licenses/>.


function onOpen(){
  
  DocumentApp.getUi().createAddonMenu()
    .addItem('Format Selection', 'formatSelection')
    .addSeparator()
    .addItem('Settings', 'showSettings')
    .addToUi();
    
  
  loadSettings();
}

function onInstall(){
  onOpen();
}


var allsettings = { fontSize:10 };

function loadSettings(){
  Logger.log('Load settings was called!');
  
  //var scriptProperties = PropertiesService.getScriptProperties();
  
  allsettings.fontSize = getScriptSettingsProperty('fontSize', 10);
  
  Logger.log('settings read');
  Logger.log(allsettings);
  
  return allsettings;
}

function getScriptSettingsProperty(keyName, defaultValue){
  var scriptProperties = PropertiesService.getScriptProperties();
  
  if (scriptProperties.getKeys().indexOf(keyName) != -1){
    var propertyValue = scriptProperties.getProperty(keyName);
    
    if (typeof(propertyValue) !== 'undefined'){
      return propertyValue;
    } else {
      return defaultValue;
    }
  } else {
    return defaultValue;
  }

}



function saveSettings(settings){
  Logger.log('saving settings');
  Logger.log(settings);
  
  var allowedKeys = ['fontSize'];
  
  if (typeof(settings) !== 'object' || settings.length == 0){
    return false;
  }
    
    var scriptProperties = PropertiesService.getScriptProperties();
    
    for(var keyName in settings){
      if (allowedKeys.indexOf(keyName) != -1) {
        scriptProperties.setProperty(keyName, settings[keyName]);
        allsettings[keyName] = settings[keyName];
      }
    }

  Logger.log('Settings saved');
  Logger.log(allsettings);
  return allsettings[keyName];
}

function formatSelection() {  
  loadSettings();
  
  var selection = DocumentApp.getActiveDocument().getSelection();
 if (selection) {
   var elements = selection.getRangeElements();
   for (var i = 0; i < elements.length; i++) {
     var element = elements[i];
     
     // Only modify elements that can be edited as text; skip images and other non-text elements.
     if (element.getElement().editAsText) {
       
       var text = element.getElement().editAsText();

       // Bold the selected part of the element, or the full element if it's completely selected.
       if (element.isPartial()) {
         //text.setBold(element.getStartOffset(), element.getEndOffsetInclusive(), true);
         text.setFontFamily(element.getStartOffset(), element.getEndOffsetInclusive(),'Courier New');
         text.setFontSize(element.getStartOffset(), element.getEndOffsetInclusive(), parseInt(allsettings.fontSize));
       } else {
         text.setFontFamily('Courier New');
         text.setFontSize(parseInt(allsettings.fontSize));
       }
     }
   }
 } else {
   // Do nothing
 }
}

function showSettings(){
  var ui = DocumentApp.getUi();
  var html = HtmlService.createTemplateFromFile('settings.html')
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(200)
      .setHeight(100);
  ui.showModalDialog(html, 'Settings');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}