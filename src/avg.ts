import * as Excel from 'exceljs'
import * as path from 'path'
import { START_ROW, GroupedRows, toDateRow, calculateAvg, writeAvgs } from './models/avg'

const fileName = process.argv[2]

if (fileName == null) {
  console.error("ERR: Please define all the params:")
  console.error("\t- path to file")
  console.error(`EXAMPLE: ${process.argv[1]} /path/to/file`)
  process.exit(9)
}

const filePath = path.parse(fileName)

loadDataFromStream(fileName)

async function loadDataFromStream(fileName: string) {
  const options: Partial<Excel.stream.xlsx.WorkbookStreamReaderOptions> = {
    sharedStrings: 'ignore',
    hyperlinks: 'ignore',
    styles: 'ignore',
    worksheets: 'emit',
    entries: 'emit',
  };
  const workbookReader = new Excel.stream.xlsx.WorkbookReader(fileName, options)
  for await (const worksheetReader of workbookReader) {
    const groupedRows: GroupedRows = new GroupedRows()
    for await (const rawRow of worksheetReader) {
      if (rawRow.number >= START_ROW) {
        const row = toDateRow(rawRow)
        groupedRows.pushFrom(row)
      }
    }
    const avgResults = calculateAvg(groupedRows)
    writeAvgs(avgResults, `${filePath.dir}/${filePath.name}_avg10min${filePath.ext}`)
  }
}
