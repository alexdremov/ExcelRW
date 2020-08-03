# ExcelRW
Lightweight Excel-based formats editor that preserves styles, macros, etc.

#### Why?
I created that package when I needed to edit .xlsm files preserving their style. Existing parse-write libraries reseted styles and views, which was unacceptable. 

## Install

```shell script
npm install @alexroar/excel-rw
```

```js
const ExcelRW = require('@alexroar/excel-rw');
```

## Usage

Before working with Excel files, they are needed to be unpacked. Therefore, part of API is asynchronous. Also, you need to specify temporary folder for packing-unpacking purposes.

**Basic usage:**
```js
let workbook = new ExcelRW(pathToTemplate, pathToTemporaryFolder)

workbook.prepareTemplate().then(async function() {
  workbook.setValue("Sheet 1", "A1", "Hello, World!") // Working with template

  await workbook.save(outputPath) // Save edits to the file

  workbook.release() // Delete temporary files from disk
})
```
