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
const StaticFeatures = require('./StaticFeatures')


/**
 * Represents single sheet in Excel file
 */
class ExcelRWorksheet {
    /**
     * Whether the sheet was altered
     */
    altered = false

    /**
     * Constructs the main object
     * @param data - parsed OpenXML /xl/sheetxx.xml contents
     * @param root - path to unpacked dir of contents
     * @param id - sheet id
     * @param full - object ExcelRW, containing current sheet
     */
    constructor(data, root, id, full) {
        this.root = root
        this.data = data
        this.id = id
        this.full = full
    }

    /**
     * Sets value for the cell
     * @param cell - target cell in format <lexical column><row number> (A1)
     * @param value - new cell value
     * @param type - auto | number | formula | text
     */
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

    /**
     * Saves sheet to the initial XML file if the sheet was changed
     *
     * @param force save even if altered === false
     */
    save(force = false) {
        if (!this.altered && !force)
            return
        let sheetInfoFile = path.join(this.root, '/xl/worksheets/', 'sheet' + this.id + '.xml')
        var builder = new xml2js.Builder();
        var xml = builder.buildObject(this.data);
        fs.writeFileSync(sheetInfoFile, xml)
    }

    /**
     * Deletes cached values in formula-cell. Use if you want formulas to be re-calculated after you open Excel file.
     */
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

    /**
     * Reads all rows in 2d array.
     * Where first index - row number +1; second index - column number
     * @param trim - trim output strings or not
     * @returns {any[][]}
     */
    readRows(trim = false) {
        const dims = this.getMaxCoordinates()
        let retData = [...Array(dims[0])].map(e => Array(dims[1]));
        if (dims[0] === 0)
            return []
        const dataInf = this.data
        if (dataInf.worksheet.sheetData[0].row === undefined)
            return retData
        for (let j = 0; j < dataInf.worksheet.sheetData[0].row.length; j++) {
            if (dataInf.worksheet.sheetData[0].row[j].c === undefined)
                continue
            for (let k = 0; k < dataInf.worksheet.sheetData[0].row[j].c.length; k++) {
                let cell = dataInf.worksheet.sheetData[0].row[j].c[k]
                let err = false
                let pos = []
                try {
                    pos = [parseInt(StaticFeatures.getRowFromCell(cell.$.r)), StaticFeatures.columnNumber(StaticFeatures.getColumnFromCell(cell.$.r))]
                    cell = this.getCellValue(cell)
                } catch (e) {
                    err = true
                }
                if (typeof cell === 'string' && trim)
                    cell = cell.trim()
                if (!err)
                    retData[pos[0] - 1][pos[1] - 1] = cell
            }
        }
        return retData
    }

    /**
     * Returns cell value
     * @param cell - target cell in format <lexical column><row number> (A1)
     * @returns {string|null}
     */
    getValue(cell) {
        const cellrow = StaticFeatures.getRowFromCell(cell)
        for (let j = 0; j < dataInf.worksheet.sheetData[0].row.length; j++) {
            if (dataInf.worksheet.sheetData[0].row[j] === undefined)
                continue
            if (dataInf.worksheet.sheetData[0].row[j].c === undefined)
                continue
            if (dataInf.worksheet.sheetData[0].row[j].$.r === undefined)
                continue
            if (parseInt(dataInf.worksheet.sheetData[0].row[j].$.r) !== cellrow)
                continue
            for (let k = 0; k < dataInf.worksheet.sheetData[0].row[j].c.length; k++) {
                let cell = dataInf.worksheet.sheetData[0].row[j].c[k]
                if (cell.$.r !== cell)
                    continue
                let err = false
                try {
                    cell = this.getCellValue(cell)
                } catch (e) {
                    err = true
                }
                if (!err)
                    return cell
            }
        }
        return null
    }

    /**
     * Get cell value
     * @param cell - parsed cell XML object.
     * @param cached - use cached sharedStrings file or load the new one
     * @returns {null|string|*}
     */
    getCellValue(cell, cached = true) {
        if (cell.$.t === undefined) {
            if (cell.v === undefined)
                return null
            return cell.v[0]
        } else if (cell.$.t === 's') {
            if (cell.v === undefined)
                return null
            return this.getTextFromShared(cell.v, cached)
        } else if (cell.f !== undefined) {
            if (cell.v === undefined)
                return null
            if (typeof cell.v[0] === 'object') {
                if (cell.v[0]._ === undefined)
                    return null
                else
                    return cell.v[0]._
            }
            return cell.v[0]
        }
        return null
    }

    /**
     * Takes shared XML information and parses it to return string value
     * @param shared - XML object of Shared String instance
     * @param cached - use cached sharedStrings file or load the new one
     * @returns {string|null|*}
     */
    getTextFromShared(shared, cached) {
        shared = this.full.readSharedStrings(cached)[shared]
        if (shared === undefined)
            return null
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

    /**
     * Get max coordinates (row and column )
     * @returns {(number)[]|number[]}
     */
    getMaxCoordinates() {
        const dims = this.data.worksheet.dimension[0].$.ref.split(':')[1]
        if (dims === undefined)
            return [0, 0]
        return [parseInt(StaticFeatures.getRowFromCell(dims)), StaticFeatures.columnNumber(StaticFeatures.getColumnFromCell(dims))]
    }
}

module.exports = ExcelRWorksheet