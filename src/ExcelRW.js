/*
 * Copyright (c) 2020.
 * Designed and developed by Aleksandr Dremov
 * dremov.me@gmail.com
 *
 */

const uuid = require('uuid')
const fs = require('fs');
const path = require('path')
const unzipper = require('unzipper')
const rimraf = require("rimraf");
const xml2js = require('xml2js')
const archiver = require('archiver');
const ExcelRWorksheet = require('./ExcelRWorksheet')
const StaticFeatures = require('./StaticFeatures')

class ExcelRW {
    unique_id = null
    tmpDir = null
    filePath = null

    sheetsIds = null
    worksheets = null

    sharedStrings = null



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

        return new Promise((resolve, reject) => {
            stream.on('close', () => {
                this.getWorksheets();
                resolve()
            });
            stream.on('error', reject);
        });
    }

    getWorksheets() {
        const shNumber = Object.keys(this.getSheetIds()).length
        this.worksheets = {}
        for (let i = 1; i <= shNumber; i++) {
            let sheetInfoFile = path.join(this.dirUnpackPath, '/xl/worksheets/', 'sheet' + i + '.xml')
            const dataInf = StaticFeatures.getXML(sheetInfoFile)
            this.worksheets[i.toString()] = new ExcelRWorksheet(dataInf, this.dirUnpackPath, i, this)
        }
    }

    getSheetIds(useCached = true) {
        if (this.sheetsIds !== null && useCached)
            return this.sheetsIds
        let sheetInfoFile = path.join(this.dirUnpackPath, '/xl/workbook.xml')

        const dataInf = StaticFeatures.getXML(sheetInfoFile)

        let sheets = {}
        for (var sh in dataInf.workbook.sheets[0].sheet) {
            sheets[dataInf.workbook.sheets[0].sheet[sh].$.name] = dataInf.workbook.sheets[0].sheet[sh].$['r:id'].slice(3)
        }
        this.sheetsIds = sheets
        return this.sheetsIds
    }

    setValue(sheet, cell, value, type = 'auto') {
        let sheetId = this.getSheetIds()[sheet]
        if (!isNaN(sheet))
            sheetId = sheet
        if (sheetId === undefined)
            throw new Error('No sheet id with identifier ' + sheet)

        return this.worksheets[sheetId.toString()].setValue(cell, value, type)
    }

    readSharedStrings(cache = false) {
        if (cache && this.sharedStrings !== null)
            return this.sharedStrings
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
        this.sharedStrings = dataInf.sst.si
        return this.sharedStrings
    }

    addToSharedStrings(string) {
        let stringsInfoFile = path.join(this.dirUnpackPath, '/xl/', 'sharedStrings.xml')
        const dataInf = StaticFeatures.getXML(stringsInfoFile)
        dataInf.sst.si.push({"t": string, $:{"xml:space": "preserve"}})
        var builder = new xml2js.Builder();
        var xml = builder.buildObject(dataInf);
        fs.writeFileSync(stringsInfoFile, xml)
        return dataInf.sst.si.length - 1
    }

    searchSharedString(string) {
        let stringsInfoFile = path.join(this.dirUnpackPath, '/xl/', 'sharedStrings.xml')
        const dataInf = StaticFeatures.getXML(stringsInfoFile)
        let data = dataInf.sst.si

        for (let i = 0; i < data.length; i++) {
            if (StaticFeatures.getTextFromSharedCell(data[i]) === string && data[i].t !== undefined) {
                return i
            }
        }
        return this.addToSharedStrings(string)
    }

    save(outputPath) {
        for (let i in this.worksheets) {
            this.worksheets[i.toString()].save()
        }

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
        for (let i in this.worksheets) {
            this.worksheets[i.toString()].deleteFormulasCache()
        }
    }
}

module.exports = ExcelRW