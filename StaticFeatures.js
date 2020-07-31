const fs = require('fs')
const xml2js = require('xml2js')

class StaticFeatures {
    static getXML(filePath) {
        const dataInf = (function () {
            let data = null, error = null;
            xml2js.parseString(fs.readFileSync(filePath), (fail, result) => {
                error = fail;
                data = result;
            });
            if (error) throw error;
            return data;
        }());
        return dataInf
    }

    static columnNumber(column) {
        var base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', i, j, result = 0;

        for (i = 0, j = column.length - 1; i < column.length; i += 1, j -= 1) {
            result += Math.pow(base.length, j) * (base.indexOf(column[i]) + 1);
        }
        return result;
    }

    static getColumnFromCell(cell) {
        return cell.match(/[a-zA-Z]+/g)[0];
    }

    static getRowFromCell(cell) {
        return cell.match(/[\d]+/g)[0];
    }
}

module.exports = StaticFeatures;