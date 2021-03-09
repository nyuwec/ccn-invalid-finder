import * as Excel from 'exceljs'
import * as path from 'path'
import { START_ROW, GroupedRows, toDateRow, aggregateCols, writeAvgs, get10MinTimeGroupKey } from './models/aggregation'

export default function avg10mFile(argv: string[]) {

  const fileName = argv[2]

  if (fileName == null) {
    console.error("ERR: Please define all the params:")
    console.error("\t- path to file")
    console.error(`EXAMPLE: ./${process.env.npm_package_name} avg /path/to/file`)
    process.exit(0)
  }

  const filePath = path.parse(fileName)

  loadDataFromStream(fileName)

  async function loadDataFromStream(fileName: string) {
    console.log(`Opening: ${fileName}`)
    const options: Partial<Excel.stream.xlsx.WorkbookStreamReaderOptions> = {
      sharedStrings: 'ignore',
      hyperlinks: 'ignore',
      styles: 'ignore',
      worksheets: 'emit',
      entries: 'emit',
    };
    const workbookReader = new Excel.stream.xlsx.WorkbookReader(fileName, options)
    for await (const worksheetReader of workbookReader) {
      const groupedRows: GroupedRows = new GroupedRows(get10MinTimeGroupKey)
      for await (const rawRow of worksheetReader) {
        if (rawRow.number >= START_ROW) {
          const row = toDateRow(rawRow)
          groupedRows.pushFrom(row)
        }
      }
      console.log(`Finished, calculating and writing AVGs...`)
      const avgResults = aggregateCols(groupedRows)
      writeAvgs(avgResults, `${filePath.dir}/${filePath.name}_avg10min${filePath.ext}`)
    }
  }
}
