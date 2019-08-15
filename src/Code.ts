function main() {
    let paramData = new ParamData("パラメーター");

    //Logger.log(paramData.colTitles);
    let params = paramData.params();
    for (let row of params) {
        Logger.log(row);
        for (let key in row) {
            Logger.log(row[key]);
        }
    }
}

class Factory {

}

class ParamData {
    sheet: GoogleAppsScript.Spreadsheet.Sheet;
    colTitles: string[];
    values: any[][];

    constructor(sheetName: string) {
        let spreadsheet = SpreadsheetApp.getActive();
        this.sheet = spreadsheet.getSheetByName(sheetName);
        let range = this.sheet.getDataRange();
        this.values = range.getValues();
        this.colTitles = this.values["0"];
    }

    params() {
        let rows=[];
        let values = this.values.slice(1);
        for (let rowIndex in values) {
            let cols = {};
            for (let colIndex in values[rowIndex]) {
                let key = this.colTitles[colIndex];
                let value = values[rowIndex][colIndex];
                cols[key] = value;
            }
            rows.push(cols);
        }
        return rows;
    }
}


function myFunction() {
    var book = SpreadsheetApp.getActive();
    var paramSheet = book.getSheetByName("パラメーター");
    var templateSheet = book.getSheetByName("パラメーター");
    // This represents ALL the data
    var range = paramSheet.getDataRange();
    var values = range.getValues();

    for (let row in values) {
        if (row == "0") {
            var cols = values[row];
        } else {
            var index = cols.indexOf("生成したファイル");
            var c = values[row][index];
            var b = "b";
        }
    }
    var a = "a";
}


