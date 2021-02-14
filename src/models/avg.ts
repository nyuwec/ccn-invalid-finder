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
    const row10MinGroupKey = this.get10MinTimeGroupKey(row.date)
    const arr = this.getOrElse(row10MinGroupKey, new Array<DateRow>())
    arr.push(row)
    this.set(row10MinGroupKey, arr)
  }

  append(groupedRow: GroupedRows) {
    groupedRow.forEach((dateRows, key) => {
      dateRows.forEach(row => {
        this.pushFrom(row)
      })
    })
  }

  private get10MinTimeGroupKey(momentVal: moment.Moment): string {
    const min = momentVal.minute()
    const groupMin = Math.floor(min / 10)
    const prefix = momentVal.format("YYYY-MM-DDTHH")
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
    vals.push(extractValue(rawRow.getCell(i)))
  }
  return {
    date: date,
    vals: vals
  }
}

export function extractValue(cell: Excel.Cell): number {
  function parseFloatFB(val: string) {
    const result = parseFloat(val)
    if (isNaN(result)) {
      return 0
    }
    return result
  }

  switch (cell.type) {
    case Excel.ValueType.Formula:
      const value = cell.value as Excel.CellFormulaValue
      return parseFloatFB(value.result as string)
    case Excel.ValueType.String:
      return parseFloatFB(eval(cell.value as string)['result'])
    case Excel.ValueType.Number:
      return parseFloatFB(cell.value as string)
    default:
      console.log(`Unknown cell.type: ${cell.type} val: ${cell.value}, row: ${cell.row}, col: ${cell.col}`)
      return parseFloatFB(cell.value as string)
  }
}

export function excelDate2Date(excelDate: number): moment.Moment {
  return moment.utc((excelDate - (25567 + 2))*86400*1000)
}

export function calculateAvg(groupedRows: GroupedRows): AvgResults {
  const avgCols = new AvgResults()
  groupedRows.forEach((groupedRows, key) => {
    if (groupedRows.length > 0) {
      const colNum = groupedRows[0].vals.length
      for (let col = 0; col < colNum; col ++) {
        const colVals = groupedRows.map(row => row.vals[col])
        const arr = avgCols.getOrElse(key, new Array<number>())
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
  const options = {
    filename: fileName,
    useStyles: false,
    useSharedStrings: false
  };

  const outXls = new Excel.stream.xlsx.WorkbookWriter(options);

  const outSheet = outXls.addWorksheet('Aggregated data');
  avgResults.forEach((avgCols, key) => {
    const groupDateShifted = moment.utc(key).add(10, 'minutes')
    const newRow = outSheet.addRow([groupDateShifted.toDate(), ...avgCols])
    newRow.commit()
  })

  outSheet.commit()
  outXls.commit()
}
