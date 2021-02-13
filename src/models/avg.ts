import { MMap } from './MMap'
import * as moment from 'moment'
import * as Excel from 'exceljs'


export const Cols = {
  "DATE": 1,
}
export const START_ROW = 5

export type DateRow = {date: moment.Moment, vals: number[]}

export class GroupedRows extends MMap<string, Array<DateRow>> {
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

export class AvgResults extends MMap<string, Array<number>> {}

// EXCEL PARSING UTILITIES
export function toDateRow(rawRow: Excel.Row): DateRow {
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

export function calculateAvg(groupedRows: GroupedRows): AvgResults {
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

export function writeAvgs(avgResults: AvgResults, fileName: string): void {
  let outXls = new Excel.Workbook();

  const outSheet = outXls.addWorksheet('Aggregated data');
  avgResults.forEach((avgCols, key) => {
    let groupDateShifted = moment.utc(key).add(10, 'minutes')
    let newRow = outSheet.addRow([groupDateShifted.toDate(), ...avgCols])
    newRow.commit()
  })

  outXls.xlsx.writeFile(fileName)
}
