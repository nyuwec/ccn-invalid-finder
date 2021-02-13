import * as Excel from 'exceljs'
import { START_ROW, DateRow, GroupedRows, AvgResults, toDateRow, calculateAvg, writeAvgs } from './models/avg'

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
