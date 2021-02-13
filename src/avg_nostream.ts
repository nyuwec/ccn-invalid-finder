import * as Excel from 'exceljs'
import { START_ROW, DateRow, GroupedRows, AvgResults, toDateRow, calculateAvg, writeAvgs } from './models/avg'

/*
These no-stream parsers are limited to heap memory.
Use the streaming version to save memory as it does not load
the whole file into memory, only line-by-line.
*/

loadDataFromCSV('data/Balatonszabadi_OPC_full.csv')
  .then((groupedRows) => {
    const avgResults = calculateAvg(groupedRows)
    writeAvgs(avgResults, 'data/new.xlsx')
  })
  .catch((err)=> {
    console.log(err)
  })


// LOADERS
async function loadDataFromXLSX(fileName: string): Promise<GroupedRows> {
  return new Excel.Workbook().xlsx.readFile(fileName)
    .then((wb) => {
      const worksheet = wb.getWorksheet(1)
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

function gatherGroupedRows(worksheet: Excel.Worksheet): GroupedRows {
  function getRowData(worksheet: Excel.Worksheet, rowNum: number): DateRow {
    const rawRow = worksheet.getRow(rowNum);
    return toDateRow(rawRow)
  }

  let groupedRows: GroupedRows = new GroupedRows()
  const lastRowNum = worksheet.actualRowCount
  for (let i=START_ROW;i<=lastRowNum;i++) {
    const row = getRowData(worksheet, i)
    groupedRows.pushFrom(row)
  }
  return groupedRows
}
