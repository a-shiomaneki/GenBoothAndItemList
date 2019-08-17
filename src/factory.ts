import { Factory, Product } from "./interface";

export class BoothAndItemSpreadsheetFactory implements Factory {
    template: { [key: string]: string };
    config: { [key: string]: string };
    macros: { [key: string]: string };
    create(config: { [key: string]: string }, template: { [key: string]: string }) {
        let product = this.createProduct(config, template);
        this.registerProduct(product);
        return product;
    }
    createProduct(config: { [key: string]: string }, template: { [key: string]: string }) {
        this.macros = {};
        this.config = config;
        this.template = template;
        this.makeMacro();
        this.evalMacro();
        let product = new BoothAndItemSpreadsheet();
        let macro: { [key: string]: string } = {};
        for (let key in this.macros) {
            macro[key] = this.macros[key];
        }
        for (let key in this.template) {
            macro[key] = this.template[key];
        }
        product.create(macro);

        return product;
    }
    registerProduct(product: Product) {
        let filename = product.filename;
        this.config["ファイル名"] = filename;
    }

    makeSheets(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {

    }
    makeMacro() {
        let year = "";
        let month = "";
        let day = "";
        let kisetsu = "";
        let startDateStr = this.config["【開催期間】開始日"];
        if (startDateStr != "") {
            let startDate = new Date(startDateStr);
            year = startDate.getFullYear().toString();
            month = startDate.getMonth().toString();
            day = startDate.getDate().toString();
        }
        let endDateStr = this.config["【開催期間】終了日"];

        this.macros["$<year>"] = year;
        this.macros["$<month>"] = month;
        let monthNumber = Number(month);
        if (monthNumber <= 2 || monthNumber >= 12) {
            kisetsu = "冬";
        } else if (monthNumber <= 5) {
            kisetsu = "春";
        } else if (monthNumber <= 8) {
            kisetsu = "夏";
        } else {
            kisetsu = "秋";
        }
        this.macros["$<kisetsu>"] = kisetsu;

        let count = this.config["イベント回数"].toString();
        this.macros["$<count>"] = count;
        this.macros["$<zenkakuCount>"] = count.replace(/[0-9]/g, function (s) {
            return String.fromCharCode(s.charCodeAt(0) + 65248);
        });
        this.macros["$<kansuujiCount>"] = count.replace(/[0-9]/g, function (s) {
            return { "0": "〇", "1": "一", "2": "二", "3": "三", "4": "四", "5": "五", "6": "六", "7": "七", "8": "八", "9": "九" }[s];
        });
    }
    evalMacro() {
        for (let keyTemplate in this.template) {
            for (let keyMacro in this.macros) {
                this.template[keyTemplate] = this.template[keyTemplate].replace(keyMacro, this.macros[keyMacro]);
            }
        }
    }
}

export class BoothAndItemSpreadsheet implements Product {
    filename: string;
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
    template: { [key: string]: string };
    create(template: { [key: string]: string }) {
        this.template = template;
        this.filename = template["$<filename>"];
        let parent = SpreadsheetApp.getActiveSpreadsheet();
        this.spreadsheet = parent.copy(this.filename);

        this.deleteConfigSheet();
        this.evalTemplateMacro();
    }
    deleteConfigSheet() {
        let sheets = this.spreadsheet.getSheets();
        for (let aSheet of sheets) {
            let sheetName = aSheet.getName();
            if (sheetName.indexOf("ctrl:") == 0) {
                this.spreadsheet.deleteSheet(aSheet);
            }
        }
    }
    evalTemplateMacro() {
        let sheets = this.spreadsheet.getSheets();
        for (let aSheet of sheets) {
            this.evalTemplateMacroToSheet(aSheet);
        }
    }
    evalTemplateMacroToSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
        let range = sheet.getDataRange();
        let values = range.getValues();
        let formulas = range.getFormulas();

        for (let keyRow in values) {
            for (let keyCol in values[keyRow]) {
                for (let macro in this.template) {
                    values[keyRow][keyCol] = values[keyRow][keyCol].replace(macro, this.template[macro]);
                }
            }
        }
        range.setValues(values);

        let rowIndex = range.getRow();
        let colIndex = range.getColumn();
        for (let keyRow in formulas) {
            for (let keyCol in formulas[keyRow]) {
                if (formulas[keyRow][keyCol] != "") {
                    let formulaRange = sheet.getRange(rowIndex + Number(keyRow), colIndex + Number(keyCol));
                    for (let macro in this.template) {
                        formulas[keyRow][keyCol]=formulas[keyRow][keyCol].replace(macro, this.template[macro]);
                    }
                    formulaRange.setFormula(formulas[keyRow][keyCol]);
                }
            }
        }

    }
}