const uuid = require('uuid')
const fs = require('fs');
const path = require('path')
const unzipper = require('unzipper')
const rimraf = require("rimraf");
const xml2js = require('xml2js')
const archiver = require('archiver');
const ExcelRWorksheet = require('./ExcelRWorksheet')

class ExcelRW {
    unique_id = null
    tmpDir = null
    filePath = null

    sheetsIds = null

    worksheets = null

    constructor(filePath, tmpDir = 'tmp') {
        this.unique_id = uuid.v4()
        this.tmpDir = tmpDir
        this.filePath = filePath
        this.dirUnpackPath = path.join(this.tmpDir, this.unique_id)
    }

    prepareTemplate() {
        const filePath = this.filePath
        const outputPath = this.dirUnpackPath
        if (!fs.existsSync(outputPath)) {
            fs.mkdirSync(outputPath);
        }
        let stream = fs.createReadStream(filePath).pipe(unzipper.Extract({path: outputPath}))

        return new Promise(function (resolve, reject) {
            stream.on('close', () => resolve());
            stream.on('error', reject);
        });
    }

    getWorksheets() {
        const shNumber = this.getSheetIds().length
    }

    getXML(filePath) {
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

    getSheetIds(useCached = true) {
        if (this.sheetsIds !== null && useCached)
            return this.sheetsIds
        let sheetInfoFile = path.join(this.dirUnpackPath, '/xl/workbook.xml')

        const dataInf = this.getXML(sheetInfoFile)

        let sheets = {}
        for (var sh in dataInf.workbook.sheets[0].sheet) {
            sheets[dataInf.workbook.sheets[0].sheet[sh].$.name] = dataInf.workbook.sheets[0].sheet[sh].$['r:id'].slice(3)
        }
        this.sheetsIds = sheets
        return this.sheetsIds
    }

    columnNumber(column) {
        var base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', i, j, result = 0;

        for (i = 0, j = column.length - 1; i < column.length; i += 1, j -= 1) {
            result += Math.pow(base.length, j) * (base.indexOf(column[i]) + 1);
        }
        return result;
    }

    getColumnFromCell(cell) {
        return cell.match(/[a-zA-Z]+/g)[0];
    }

    getRowFromCell(cell) {
        return cell.match(/[\d]+/g)[0];
    }

    setValue(sheet, cell, value, type = 'auto') {
        let sheetId = this.getSheetIds()[sheet]
        if (!isNaN(sheet))
            sheetId = sheet
        if (sheetId === undefined)
            throw new Error('No sheet id with identifier ' + sheet)

        let sheetInfoFile = path.join(this.dirUnpackPath, '/xl/worksheets/', 'sheet' + sheetId + '.xml')
        const dataInf = this.getXML(sheetInfoFile)
        const cellRow = this.getRowFromCell(cell)
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
            dataInf.worksheet.sheetData[0].row[foundRow].c[foundCell].v = this.searchSharedString(value)
        } else if (type === 'formula') {
            delete dataInf.worksheet.sheetData[0].row[foundRow].c[foundCell].$.t
            delete dataInf.worksheet.sheetData[0].row[foundRow].c[foundCell].v
            dataInf.worksheet.sheetData[0].row[foundRow].c[foundCell].f = value
        }

        var builder = new xml2js.Builder();
        var xml = builder.buildObject(dataInf);
        fs.writeFileSync(sheetInfoFile, xml)
    }

    readSharedStrings() {
        let stringsInfoFile = path.join(this.dirUnpackPath, '/xl/', 'sharedStrings.xml')
        const dataInf = (function () {
            let data = null, error = null;
            xml2js.parseString(fs.readFileSync(stringsInfoFile), (fail, result) => {
                error = fail;
                data = result;
            });
            if (error) throw error;
            return data;
        }());

        return dataInf.sst.si
    }

    addToSharedStrings(string) {
        let stringsInfoFile = path.join(this.dirUnpackPath, '/xl/', 'sharedStrings.xml')
        const dataInf = this.getXML(stringsInfoFile)
        dataInf.sst.si.push({"t": string})
        var builder = new xml2js.Builder();
        var xml = builder.buildObject(dataInf);
        fs.writeFileSync(stringsInfoFile, xml)
        return dataInf.sst.si.length - 1
    }

    searchSharedString(string) {
        let stringsInfoFile = path.join(this.dirUnpackPath, '/xl/', 'sharedStrings.xml')
        const dataInf = this.getXML(stringsInfoFile)
        let data = dataInf.sst.si

        for (let i = 0; i < data.length; i++) {
            if (data[i].t === string && data[i].t !== undefined) {
                return i
            }
        }
        return this.addToSharedStrings(string)
    }

    save(outputPath) {
        const unpackedDir = this.dirUnpackPath
        const archive = archiver('zip');
        const stream = fs.createWriteStream(outputPath);

        return new Promise((resolve, reject) => {
            archive
                .directory(unpackedDir, false)
                .on('error', err => reject(err))
                .pipe(stream)
            ;

            stream.on('close', () => resolve());
            archive.finalize();
        });
    }

    release(cb = function () {

    }) {
        if (fs.existsSync(this.dirUnpackPath))
            rimraf(path.join(this.dirUnpackPath), [], cb)
    }

    deleteFormulasCache() {
        let sheetIds = this.getSheetIds()
        for (let id in sheetIds) {
            let sheetInfoFile = path.join(this.dirUnpackPath, '/xl/worksheets/', 'sheet' + sheetIds[id] + '.xml')
            const dataInf = this.getXML(sheetInfoFile)
            if (dataInf.worksheet.sheetData[0].row === undefined)
                continue
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
            var builder = new xml2js.Builder();
            var xml = builder.buildObject(dataInf);
            fs.writeFileSync(sheetInfoFile, xml)
        }
    }
}

module.exports = ExcelRW