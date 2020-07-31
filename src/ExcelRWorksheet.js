const uuid = require('uuid')
const fs = require('fs');
const path = require('path')
const unzipper = require('unzipper')
const rimraf = require("rimraf");
const xml2js = require('xml2js')
const archiver = require('archiver');
const StaticFeatures = require('./StaticFeatures')

class ExcelRWorksheet {

    altered = false

    constructor(data, root, id, full) {
        this.root = root
        this.data = data
        this.id = id
        this.full = full
    }

    setValue(cell, value, type = 'auto') {
        this.altered = true
        const dataInf = this.data
        const cellRow = StaticFeatures.getRowFromCell(cell)
        let curPos = dataInf.worksheet.sheetData[0].row

        let foundRow = -1

        for (let i = 0; i < curPos.length; i++) {
            if (curPos[i].$.r === cellRow) {
                foundRow = i;
                break
            }
        }
        if (foundRow === -1) {
            dataInf.worksheet.sheetData[0].row.push({
                $: {
                    r: cellRow,
                },
                c: []
            })
            foundRow = dataInf.worksheet.sheetData[0].row.length - 1
        }
        curPos = dataInf.worksheet.sheetData[0].row

        let foundCell = -1

        curPos = curPos[foundRow].c

        for (let i = 0; i < curPos.length; i++) {
            if (curPos[i].$.r === cell) {
                foundCell = i;
                break
            }
        }

        if (foundCell === -1) {
            dataInf.worksheet.sheetData[0].row[foundRow].c.push({
                $: {
                    r: cell,
                }
            })
            foundCell = dataInf.worksheet.sheetData[0].row[foundRow].c.length - 1
        }

        if ((!isNaN(value) && /^\d+$/.test(value) || typeof value == "number") && (type === 'auto' || type === 'number')) {
            delete dataInf.worksheet.sheetData[0].row[foundRow].c[foundCell].$.t
            dataInf.worksheet.sheetData[0].row[foundRow].c[foundCell].v = value
        } else if (type === 'auto' || type === 'text') {
            dataInf.worksheet.sheetData[0].row[foundRow].c[foundCell].$.t = "s"
            dataInf.worksheet.sheetData[0].row[foundRow].c[foundCell].v = this.full.searchSharedString(value)
        } else if (type === 'formula') {
            delete dataInf.worksheet.sheetData[0].row[foundRow].c[foundCell].$.t
            delete dataInf.worksheet.sheetData[0].row[foundRow].c[foundCell].v
            dataInf.worksheet.sheetData[0].row[foundRow].c[foundCell].f = value
        }
        this.data = dataInf
    }

    save(force = false) {
        if (!this.altered && !force)
            return
        let sheetInfoFile = path.join(this.root, '/xl/worksheets/', 'sheet' + this.id + '.xml')
        var builder = new xml2js.Builder();
        var xml = builder.buildObject(this.data);
        fs.writeFileSync(sheetInfoFile, xml)
    }

    deleteFormulasCache() {
        this.altered = true
        const dataInf = this.data
        if (dataInf.worksheet.sheetData[0].row === undefined)
            return
        for (let j = 0; j < dataInf.worksheet.sheetData[0].row.length; j++) {
            if (dataInf.worksheet.sheetData[0].row[j].c === undefined)
                continue
            for (let k = 0; k < dataInf.worksheet.sheetData[0].row[j].c.length; k++) {
                let cell = dataInf.worksheet.sheetData[0].row[j].c[k]
                if (cell.f !== undefined) {
                    delete dataInf.worksheet.sheetData[0].row[j].c[k].v
                }
            }
        }
        this.data = dataInf
    }
}

module.exports = ExcelRWorksheet