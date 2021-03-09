import * as Excel from 'exceljs'
import * as path from 'path'
import { GroupedRows, toDateRow, aggregateCols, writeAvgs, get1MinTimeGroupKey, sumValues, avgValues, AggregateFunction } from './models/aggregation'

export default function avgFile1min(argv: string[]) {
  const START_ROW = 2
  const fileName = argv[2]

  if (fileName == null) {
    console.error("ERR: Please define all the params:")
    console.error("\t- path to file")
    console.error(`EXAMPLE: ${process.title} avg1m /path/to/file`)
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
      const groupedRows: GroupedRows = new GroupedRows(get1MinTimeGroupKey)
      for await (const rawRow of worksheetReader) {
        if (rawRow.number >= START_ROW) {
          try {
            const row = toDateRow(rawRow, 3, false)
            groupedRows.pushFrom(row)
          } catch (e) {
            console.log(`Error parsing row #${rawRow.number}: ${e}. SKIPPED`)
          }
        }
      }
      console.log(`Finished, calculating and writing AVGs...`)
      const agrSetup: AggregateFunction[] = [
        avgValues, sumValues, sumValues, avgValues, avgValues, avgValues,
        sumValues, sumValues, sumValues, sumValues, sumValues, sumValues, sumValues, sumValues, sumValues, sumValues,
        sumValues, sumValues, sumValues, sumValues, sumValues, sumValues, sumValues, sumValues, sumValues, sumValues,
        sumValues, sumValues, sumValues, sumValues, sumValues, sumValues, sumValues, sumValues, sumValues, sumValues
      ]

      const avgResults = aggregateCols(groupedRows, agrSetup)
      writeAvgs(avgResults,
        `${filePath.dir}/${filePath.name}_avg1min${filePath.ext}`,
        1,
        [['Date', 'T Ambient (C)', 'Number Conc (#/cm^3)', 'LWC (g/m^3)', 'MVD (um)', 'ED (um)', 'Applied PAS (m/s)', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '16', '18', '20', '22', '24', '26', '28', '30', '32', '34', '36', '38', '40', '42', '44', '46', '48', '50']]
      )
    }
  }
}
