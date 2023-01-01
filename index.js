const fs = require('fs')
const Tesseract = require('tesseract.js')
const ExcelJS = require('exceljs')

// path to the folder containing the images
const folderPath = './Process'

async function extractTextFromImages() {
  try {
    console.log('Reading files from folder: ' + folderPath)

    // read the contents of the directory
    const files = await fs.promises.readdir(folderPath)

    console.log('Files in folder: ' + files.length)

    let workbook
    let sheet

    // check if the output file exists
    try {
      // read the workbook from the file if it exists
      workbook = new ExcelJS.Workbook()
      await workbook.xlsx.readFile('output.xlsx')
      sheet = workbook.getWorksheet('Output')
    } catch (error) {
      // create a new workbook and add a sheet if the file does not exist
      workbook = new ExcelJS.Workbook()
      sheet = workbook.addWorksheet('Output')
    }

    // loop through the files
    for (const file of files) {
      let rowExists = false

      // loop through the rows in the sheet and check the value in the first column
      sheet.eachRow((row) => {
        if (row.getCell(1).value === file) {
          console.log(`Skipping ${file} because it already exists in the sheet`)
          rowExists = true
        }
      })

      if (rowExists) {
        continue
      }

      // read the file
      const data = await fs.promises.readFile(`${folderPath}/${file}`)

      // extract the text from the image using Tesseract.js
      const result = await Tesseract.recognize(data)
      console.log(result.data.text)

      // add a new row to the sheet with the extracted text
      console.log(`Adding ${file} to the sheet`)
      sheet.addRow([file, result.data.text])

      // write the workbook to disk
      console.log('Writing to output.xlsx')
      await workbook.xlsx.writeFile('output.xlsx')
    }
  } catch (err) {
    console.error(err)
  }
}

extractTextFromImages()
