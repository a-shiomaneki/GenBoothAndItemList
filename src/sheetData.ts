export class SheetAsDatabase {
    sheet: GoogleAppsScript.Spreadsheet.Sheet;
    range: GoogleAppsScript.Spreadsheet.Range;
    colTitles: string[];
    values: any[][];
    sheetName: string;

    constructor(sheetName: string) {
        this.sheetName = sheetName;
    }

    readData() {
        let spreadsheet = SpreadsheetApp.getActive();
        this.sheet = spreadsheet.getSheetByName(this.sheetName);
        this.range = this.sheet.getDataRange();
        this.values = this.range.getValues();
        this.colTitles = this.values["0"];
    }
    writeData() {
        this.range.setValues(this.values);
        SpreadsheetApp.flush();
    }
}


export class ConfigData extends SheetAsDatabase {
    configs() {
        let rows = [];
        let values = this.values.slice(1);
        for (let row of values) {
            let cols = {};
            for (let colIndex in row) {
                let key = this.colTitles[colIndex];
                let value = row[colIndex];
                cols[key] = value;
            }
            rows.push(cols);
        }
        return rows;
    }
    setConfig(indexRow: number, updateRow: {}) {
        for (let keyCol in this.colTitles) {
            let title = this.colTitles[Number(keyCol)];
            this.values[indexRow + 1][Number(keyCol)] = updateRow[title];
        }
        this.writeData();
    }
}

export class MacroData extends SheetAsDatabase {
    templates() {
        let rows = {};
        let values = this.values.slice(1);
        for (let row of values) {
            let eventType: string;
            let cols = {};
            for (let colIndex in row) {
                let key = this.colTitles[colIndex];
                let value = row[colIndex];
                if (key == "イベントタイプ") {
                    eventType = value;
                } else {
                    cols[key] = value;
                }
            }
            rows[eventType] = cols;
        }
        return rows;
    }
}