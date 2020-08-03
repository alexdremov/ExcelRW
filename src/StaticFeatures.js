/*
 * Copyright (c) 2020.
 * Designed and developed by Aleksandr Dremov
 * dremov.me@gmail.com
 *
 */

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

    static getTextFromSharedCell(shared) {
        if (shared.t !== undefined) {
            if (typeof shared.t[0] === 'string')
                return shared.t[0]
            if (shared.t[0]._ !== undefined)
                return shared.t[0]._
            return null
        } else if (shared.r !== undefined) {
            let outStr = ''
            for (let i in shared.r) {
                if (shared.r[i].t === undefined)
                    continue
                if (typeof shared.r[i].t === 'string')
                    outStr += shared.r[i].t
                else if (typeof shared.r[i]._ === 'string')
                    outStr += shared.r[i]._
            }
            return outStr
        }
        return null
    }
}

module.exports = StaticFeatures;