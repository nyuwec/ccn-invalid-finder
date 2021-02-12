import * as Excel from 'exceljs'
import * as moment from 'moment'
import { MMap } from './models/MMap'

let inXls = new Excel.Workbook();
let outXls = new Excel.Workbook();

const Cols = {
  "DATE": 1,
  "LAST": 31
}
const START_ROW = 5

type RowType = {date: string, vals: number[]}
class GroupedRows extends MMap<string, Array<RowType>> {}

loadDataFromCSV('data/Balatonszabadi_OPC_full.csv')
  .then((groupedRows) => {
    calculateAndWrite(groupedRows)
  })
  .catch((err)=> {
    console.log(err)
  })

async function loadDataFromCSV(fileName: string): Promise<GroupedRows> {
  return inXls.csv.readFile(fileName)
    .then((wb) => {
      // let worksheet = wb.getWorksheet(1);
      let worksheet = wb

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
  )
}

function calculateAndWrite(groupedRows: GroupedRows): void {
  const outSheet = outXls.addWorksheet('Aggregated data');
  let avgCols = new MMap<string, Array<number>>()
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
    let groupDateShifted = moment.utc(key).add(10, 'minutes')
    let newRow = outSheet.addRow([groupDateShifted.toDate(), ...avgCols.getOrElse(key, [])])
    newRow.commit()
  })
  outXls.xlsx.writeFile('data/new.xlsx')
}

function getRowData(worksheet: Excel.Worksheet, rowNum: number): RowType {
  let row = worksheet.getRow(rowNum);
  let date: Excel.CellValue = row.getCell(Cols.DATE).value
  let vals: number[] = []
  for (let i=Cols.DATE+1;i<=Cols.LAST;i++) {
   vals.push(row.getCell(i).value as number)
  }
  return {
    date: date as string,
    vals: vals
  }
}

function get10MinTimeGroupKey(timeString: string): string {
  let ts = moment.utc(timeString, "YYYY.MM.DD hh:mm:ss")
  let min = ts.minute()
  let groupMin = Math.floor(min / 10)
  let prefix = ts.format("YYYY-MM-DDTHH")
  return `${prefix}:${groupMin}0:00.000Z`
}
