import * as Excel from 'exceljs'
import * as moment from 'moment'
import { MMap } from './models/MMap'

let outXls = new Excel.Workbook();

const Cols = {
  "DATE": 1,
  "LAST": 31
}
const START_ROW = 5

type RowType = {date: moment.Moment, vals: number[]}
class GroupedRows extends MMap<string, Array<RowType>> {}
class AvgResults extends MMap<string, Array<number>> {}

// loadDataFromStream('data/Balatonszabadi_OPC_full.xlsx')

loadDataFromCSV('data/Balatonszabadi_OPC_full.csv')
  .then((groupedRows) => {
    const avgResults = calculateAvg(groupedRows)
    writeAvgs(avgResults)
  })
  .catch((err)=> {
    console.log(err)
  })


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
    sharedStrings: 'emit',
    hyperlinks: 'emit',
    worksheets: 'emit',
    entries: 'emit',
  };
  const workbookReader = new Excel.stream.xlsx.WorkbookReader(fileName, options)
  for await (const worksheetReader of workbookReader) {
    let groupedRows: GroupedRows = new GroupedRows()
    for await (const rawRow of worksheetReader) {
      if (rawRow.number >= START_ROW) {
        let row = getRawRowData(rawRow)
        let row10MinGroupKey = get10MinTimeGroupKey(row.date)
        let arr = groupedRows.getOrElse(row10MinGroupKey, new Array<RowType>())
        arr.push(row)
        groupedRows.set(row10MinGroupKey, arr)
      }
    }
    const avgResults = calculateAvg(groupedRows)
    writeAvgs(avgResults)
  }
  return null
}


// AGGREGATOR FUNCTIONS
function gatherGroupedRows(worksheet: Excel.Worksheet): GroupedRows {
  let groupedRows: GroupedRows = new GroupedRows()
  const lastRowNum = worksheet.actualRowCount
  for (let i=START_ROW;i<=lastRowNum;i++) {
    let row = getRowData(worksheet, i)
    let row10MinGroupKey = get10MinTimeGroupKey(row.date)
    let arr = groupedRows.getOrElse(row10MinGroupKey, new Array<RowType>())
    arr.push(row)
    groupedRows.set(row10MinGroupKey, arr)
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
function getRowData(worksheet: Excel.Worksheet, rowNum: number): RowType {
  let rawRow = worksheet.getRow(rowNum);
  return getRawRowData(rawRow)
}

function getRawRowData(rawRow: Excel.Row): RowType {
  const dateCell = rawRow.getCell(Cols.DATE)
  let date: moment.Moment
  if (dateCell.type == Excel.ValueType.Number) {
    date = excelDate2Date(dateCell.value as number)
  } else {
    date = moment.utc(dateCell.value as string, "YYYY.MM.DD hh:mm:ss")
  }
  let vals: number[] = []
  for (let i=Cols.DATE+1;i<=Cols.LAST;i++) {
   vals.push(rawRow.getCell(i).value as number)
  }
  return {
    date: date,
    vals: vals
  }
}

function get10MinTimeGroupKey(momentVal: moment.Moment): string {
  let min = momentVal.minute()
  let groupMin = Math.floor(min / 10)
  let prefix = momentVal.format("YYYY-MM-DDTHH")
  return `${prefix}:${groupMin}0:00.000Z`
}

function excelDate2Date(excelDate: number): moment.Moment {
  return moment.utc((excelDate - (25567 + 2))*86400*1000)
}
