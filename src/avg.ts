import * as Excel from 'exceljs'
import * as moment from 'moment'
import { MMap } from './models/MMap'

let workbook = new Excel.Workbook();

const Cols = {
  "DATE": 1,
  "LAST": 31
}
const START_ROW = 5

type RowType = {date: string, vals: number[]}

workbook.xlsx.readFile('data/2.xlsx')
  .then((wb) => {
    let worksheet = wb.getWorksheet(1);

    let groupedRows = new MMap<string, Array<RowType>>()
    for (let i=START_ROW;i<=worksheet.actualRowCount;i++) {
      let row = getRowData(worksheet, i)
      let row10MinGroup = get10MinTimeGroupKey(row.date as string)
      let arr = groupedRows.getOrElse(row10MinGroup, new Array<RowType>())
      arr.push(row)
      groupedRows.set(row10MinGroup, arr)
    }
    groupedRows.forEach((groupedRows, index) => {
      console.log(`"${index}" has ${groupedRows.length} items`)
      let avgCols = []
      if (groupedRows.length > 0) {
        let colNum = groupedRows[0].vals.length
        for (let col = 0; col < colNum; col ++) {
          let colVals = groupedRows.map(row => row.vals[col])
          avgCols[col] = (
            (colVals.reduce((prev, current) => prev + current, 0) / colVals.length)
          ).toFixed(4)
        }
      }
      console.log(`    AVGS: ${avgCols}`)
    })
    //row.getCell(1).value = 5; // A5's value set to 5
    //row.commit();
    //return workbook.xlsx.writeFile('new.xlsx');
    })
    .catch((err)=> {
      console.log(err)
    })

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
  let ts = moment.utc(timeString)
  let min = ts.minute()
  let groupMin = Math.floor(min / 10)
  let prefix = ts.format("YYYY-MM-DD:HH")
  return `${prefix}:${groupMin}`
}
