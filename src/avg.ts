import * as Excel from 'exceljs'
import * as moment from 'moment'
import { MMap } from './models/MMap'

let outXls = new Excel.Workbook();

const Cols = {
  "DATE": 1,
}
const START_ROW = 5

type DateRow = {date: moment.Moment, vals: number[]}
class GroupedRows extends MMap<string, Array<DateRow>> {
  pushFrom(row: DateRow) {
    let row10MinGroupKey = this.get10MinTimeGroupKey(row.date)
    let arr = this.getOrElse(row10MinGroupKey, new Array<DateRow>())
    arr.push(row)
    this.set(row10MinGroupKey, arr)
  }

  private get10MinTimeGroupKey(momentVal: moment.Moment): string {
    let min = momentVal.minute()
    let groupMin = Math.floor(min / 10)
    let prefix = momentVal.format("YYYY-MM-DDTHH")
    return `${prefix}:${groupMin}0:00.000Z`
  }
}
class AvgResults extends MMap<string, Array<number>> {}

loadDataFromStream('data/Balatonszabadi_OPC_full.xlsx')

// loadDataFromCSV('data/Balatonszabadi_OPC_full.csv')
//   .then((groupedRows) => {
//     const avgResults = calculateAvg(groupedRows)
//     writeAvgs(avgResults)
//   })
//   .catch((err)=> {
//     console.log(err)
//   })


// LOADERS
async function loadDataFromXLSX(fileName: string): Promise<GroupedRows> {
  return new Excel.Workbook().xlsx.readFile(fileName)
    .then((wb) => {
      let worksheet = wb.getWorksheet(1)
      return gatherGroupedRows(worksheet)
    })
}

async function loadDataFromCSV(fileName: string): Promise<GroupedRows> {
  return new Excel.Workbook().csv.readFile(fileName)
    .then((worksheet) => {
      return gatherGroupedRows(worksheet)
    }
  )
}

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
    let groupedRows: GroupedRows = new GroupedRows()
    for await (const rawRow of worksheetReader) {
      if (rawRow.number >= START_ROW) {
        let row = toDateRow(rawRow)
        groupedRows.pushFrom(row)
      }
    }
    const avgResults = calculateAvg(groupedRows)
    writeAvgs(avgResults)
  }
}


// AGGREGATOR FUNCTIONS
function gatherGroupedRows(worksheet: Excel.Worksheet): GroupedRows {
  function getRowData(worksheet: Excel.Worksheet, rowNum: number): DateRow {
    let rawRow = worksheet.getRow(rowNum);
    return toDateRow(rawRow)
  }

  let groupedRows: GroupedRows = new GroupedRows()
  const lastRowNum = worksheet.actualRowCount
  for (let i=START_ROW;i<=lastRowNum;i++) {
    let row = getRowData(worksheet, i)
    groupedRows.pushFrom(row)
  }
  return groupedRows
}

function writeAvgs(avgResults: AvgResults): void {
  const outSheet = outXls.addWorksheet('Aggregated data');
  avgResults.forEach((avgCols, key) => {
    let groupDateShifted = moment.utc(key).add(10, 'minutes')
    let newRow = outSheet.addRow([groupDateShifted.toDate(), ...avgCols])
    newRow.commit()
  })
  outXls.xlsx.writeFile('data/new.xlsx')
}

function calculateAvg(groupedRows: GroupedRows): AvgResults {
  let avgCols = new AvgResults()
  groupedRows.forEach((groupedRows, key) => {
    if (groupedRows.length > 0) {
      let colNum = groupedRows[0].vals.length
      for (let col = 0; col < colNum; col ++) {
        let colVals = groupedRows.map(row => row.vals[col])
        let arr = avgCols.getOrElse(key, new Array<number>())
        arr[col] = (
          colVals.reduce(
            (prev, current) => prev + current,
            0
          ) / colVals.length)
        avgCols.set(key, arr)
      }
    }
  })
  return avgCols
}


// EXCEL PARSING UTILITIES
function toDateRow(rawRow: Excel.Row): DateRow {
  const dateCell = rawRow.getCell(Cols.DATE)
  const date: moment.Moment = function () {
    if (dateCell.type == Excel.ValueType.Number) {
      return excelDate2Date(dateCell.value as number)
    } else {
      return moment.utc(dateCell.value as string, "YYYY.MM.DD hh:mm:ss")
    }
  }()

  const vals: number[] = Array<number>()
  for (let i=Cols.DATE+1;i<=rawRow.cellCount;i++) {
   vals.push(rawRow.getCell(i).value as number)
  }
  return {
    date: date,
    vals: vals
  }
}

function excelDate2Date(excelDate: number): moment.Moment {
  return moment.utc((excelDate - (25567 + 2))*86400*1000)
}
