import * as moment from 'moment'
import * as Excel from 'exceljs'

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
  return moment.utc(Math.round((excelDate - 25569.0)*86400000.0))
}
