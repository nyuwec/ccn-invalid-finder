import { MMap } from './MMap'
import * as moment from 'moment'
import * as Excel from 'exceljs'


export const START_ROW = 5

export type DateRow = {date: moment.Moment, vals: number[]}

export type GroupKeyFunction = (m: moment.Moment) => string;

export class GroupedRows extends MMap<string, Array<DateRow>> {
  private groupKeyGen: GroupKeyFunction

  constructor (groupKeyGen: GroupKeyFunction) {
    super()
    this.groupKeyGen = groupKeyGen
  }

  pushFrom(row: DateRow) {
    const rowGroupKey = this.groupKeyGen(row.date)
    const arr = this.getOrElse(rowGroupKey, new Array<DateRow>())
    arr.push(row)
    this.set(rowGroupKey, arr)
  }

  append(groupedRow: GroupedRows) {
    groupedRow.forEach(dateRows => {
      dateRows.forEach(row => {
        this.pushFrom(row)
      })
    })
  }
}

export function get10MinTimeGroupKey(momentVal: moment.Moment): string {
  const min = momentVal.minute()
  const groupMin = Math.floor(min / 10)
  const prefix = momentVal.format("YYYY-MM-DDTHH")
  return `${prefix}:${groupMin}0:00.000Z`
}

export function get1MinTimeGroupKey(momentVal: moment.Moment): string {
  const min = momentVal.minute().toString().padStart(2, '0')
  const prefix = momentVal.format("YYYY-MM-DDTHH")
  return `${prefix}:${min}:00.000Z`
}

export class AggregateResults extends MMap<string, Array<number>> {}

// EXCEL PARSING UTILITIES
export function toDateRow(rawRow: Excel.Row, dateColPos: number = 1, fallback: boolean = true): DateRow {
  const cell = rawRow.getCell(dateColPos)
  const date: moment.Moment = toMoment(cell)
  // console.log(`dateCell: ${cell}, type: ${cell.type}, parsed: ${date}`)

  const vals: number[] = Array<number>()
  for (let i=dateColPos+1;i<=rawRow.cellCount;i++) {
    vals.push(extractNumber(rawRow.getCell(i), fallback))
  }
  return {
    date: date,
    vals: vals
  }
}

export function toMoment(dateCell: Excel.Cell): moment.Moment {
  const dateNum = extractNumber(dateCell)
  if (dateNum == 0) {
    return moment.utc(dateCell.value as string, "YYYY.MM.DD hh:mm:ss")
  } else {
    return excelDate2Date(dateNum)
  }
}

class NotANumberError extends Error {
  constructor(message: string) {
    super(message)
    this.name = this.constructor.name
  }
}

export function extractNumber(cell: Excel.Cell, fallback: boolean = true): number {
  function parseFloatFB(val: string) {
    const result = parseFloat(val)
    if (isNaN(result)) {
      if (fallback) {
        return 0
      } else {
        throw new NotANumberError(`wrong number: ${val}`)
      }
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

export type AggregateFunction = (vals: number[]) => number

export function aggregateCols(groupedRows: GroupedRows, setup: AggregateFunction[] = []): AggregateResults {
  const agrCols = new AggregateResults()
  groupedRows.forEach((groupedRows, key) => {
    if (groupedRows.length > 0) {
      const colNum = groupedRows[0].vals.length
      for (let col = 0; col < colNum; col ++) {
        const colVals = groupedRows.map(row => row.vals[col])
        const arr = agrCols.getOrElse(key, new Array<number>())
        if (setup[col]) {
          arr[col] = setup[col](colVals)
        } else {
          arr[col] = avgValues(colVals)
        }
        agrCols.set(key, arr)
      }
    }
  })
  return agrCols
}

export function avgValues(vals: number[]): number {
  return sumValues(vals) / vals.length
}

export function sumValues(vals: number[]): number {
  return vals.reduce(
    (prev, current) => prev + current,
    0
  )
}

export function writeAvgs(
  agrResults: AggregateResults,
  fileName: string,
  shiftDateByMin = 10,
  headerRows: string[][] = [[]]
): void {
  const options = {
    filename: fileName,
    useStyles: false,
    useSharedStrings: false
  };

  const outXls = new Excel.stream.xlsx.WorkbookWriter(options);

  const outSheet = outXls.addWorksheet('Aggregated data');
  if (headerRows.length > 0) {
    headerRows.forEach((headerRow) => {
      const row = outSheet.addRow(headerRow)
      row.eachCell(cell => {
        cell.font = {
          bold: true,
          color: { argb: 'FF4472c4'}
        }
      })
      row.commit()
    })
  }
  agrResults.forEach((avgCols, key) => {
    const groupDateShifted = moment.utc(key).add(shiftDateByMin, 'minutes')
    const newRow = outSheet.addRow([groupDateShifted.toDate(), ...avgCols])
    newRow.commit()
  })

  outSheet.commit()
  outXls.commit()
}
