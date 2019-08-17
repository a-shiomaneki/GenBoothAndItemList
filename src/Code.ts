import { ConfigData, MacroData } from "./sheetData";
import { BoothAndItemSpreadsheetFactory } from "./factory";

function main() {
    let configData = new ConfigData("ctrl:設定");
    configData.readData();

    //Logger.log(configData.colTitles);
    let configs = configData.configs();
    for (let thisConfig of configs) {
        Logger.log(thisConfig);
        for (let key in thisConfig) {
            Logger.log(thisConfig[key]);
        }
    }

    let templateData = new MacroData("ctrl:マクロ");
    templateData.readData();
    let templates = templateData.templates();
    for (let eventType in templates) {
        for (let key in templates[eventType]) {
            Logger.log(templates[eventType][key]);
        }
    }

    for (let keyRow in configs) {
        let thisConfig = configs[keyRow];
        if (thisConfig["更新?"]) {
            let thisTemplate = templates[thisConfig["イベントタイプ"]];
            let factory = new BoothAndItemSpreadsheetFactory();
            let spreadsheet = factory.create(thisConfig, thisTemplate);
            configData.setConfig(Number(keyRow), factory.config);

            let folderName = thisConfig["フォルダ"];
            if (folderName != "") {
                let folder: GoogleAppsScript.Drive.Folder;
                let folderItr = DriveApp.getFoldersByName(folderName);
                if (folderItr.hasNext()) {
                    folder = folderItr.next();
                } else {
                    folder = DriveApp.createFolder(folderName);
                }
                let file = DriveApp.getFileById(spreadsheet.spreadsheetId);
                let parentsItr = file.getParents();

                while (parentsItr.hasNext()) {
                    let parent = parentsItr.next();
                    parent.removeFile(file);
                }
                folder.addFile(file);
                file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
            }
        }
    }
}
