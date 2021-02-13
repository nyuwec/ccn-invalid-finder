import * as Excel from 'exceljs'
import * as moment from 'moment'
import { Cols, START_ROW, DateRow, GroupedRows, AvgResults, toDateRow, calculateAvg, writeAvgs } from './models/avg'

loadDataFromStream('data/Balatonszabadi_OPC_full.xlsx')

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
    writeAvgs(avgResults, 'data/new.xlsx')
  }
}
