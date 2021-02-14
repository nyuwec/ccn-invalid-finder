import * as fs from 'fs'
import * as path from 'path'
import * as Excel from 'exceljs'
import * as moment from 'moment'
import { DateRow, GroupedRows, AvgResults, excelDate2Date, calculateAvg, writeAvgs, extractValue } from './models/avg'

const START_ROW = 9
const BIN_FIRST_COL = 47
const BIN_LAST_COL = 71
const PM_FIRST_COL = 41
const PM_LAST_COL = 46

loadDataFromFolder('./data_preprocessed/')
  .then(resultedGRows => {
    const groupedRows: GroupedRows = new GroupedRows()
    resultedGRows.forEach(gr => {
      groupedRows.append(gr)
    })
    writeGroupedRows(groupedRows, 'data_preprocessed/full.xlsx')

    const avgResult = calculateAvg(groupedRows)
    writeAvgs(avgResult, 'data_preprocessed/full_avg.xlsx')
  })

async function loadDataFromFolder(dirName: string): Promise<GroupedRows[]> {
  let allFiles: fs.Dirent[] = fs.readdirSync(dirName, {
    encoding: 'utf8',
    withFileTypes: true
  })

  const all = allFiles
    .filter((dirEntity) => {
      return (dirEntity.isFile() && path.extname(dirEntity.name) == '.xlsx')
    })
    .map(dirEntity => dirName + dirEntity.name)
    .sort()
    .map(fileName => {
      const rows = readFile(fileName)
      return rows
    })

  return Promise.all(all)
}

async function readFile(fileName: string): Promise<GroupedRows> {
  const options: Partial<Excel.stream.xlsx.WorkbookStreamReaderOptions> = {
    sharedStrings: 'ignore',
    hyperlinks: 'ignore',
    styles: 'ignore',
    worksheets: 'emit',
    entries: 'emit',
  }
  const groupedRows: GroupedRows = new GroupedRows()

  console.log(`Open ${fileName}`)
  const workbookReader = new Excel.stream.xlsx.WorkbookReader(fileName, options)
  for await (const worksheetReader of workbookReader) {
    for await (const rawRow of worksheetReader) {
      if (rawRow.number >= START_ROW
        && rawRow.cellCount >= BIN_LAST_COL
        && rawRow.getCell(1).type != Excel.ValueType.Null
      ) {
        const row = toDateRow(rawRow)
        groupedRows.pushFrom(row)
      }
    }
  }
  console.log(`Finished ${fileName}`)
  return groupedRows
}

function toDateRow(rawRow: Excel.Row): DateRow {
  const dateCell = rawRow.getCell(1)
  const date: moment.Moment = function () {
    if (dateCell.type == Excel.ValueType.Number) {
      return excelDate2Date(dateCell.value as number)
    } else {
      return moment.utc(dateCell.value as string, "YYYY.MM.DD hh:mm:ss")
    }
  }()

  const vals: number[] = Array<number>()
  for (let i=BIN_FIRST_COL;i<=BIN_LAST_COL;i++) {
    vals.push(extractValue(rawRow.getCell(i)))
  }
  for (let i=PM_FIRST_COL;i<=PM_LAST_COL;i++) {
    vals.push(extractValue(rawRow.getCell(i)))
  }
  return {
    date: date,
    vals: vals
  }
}

function writeGroupedRows(groupedRows: GroupedRows, fileName: string) {
  const options = {
    filename: fileName,
    useStyles: false,
    useSharedStrings: false
  };

  const outXls = new Excel.stream.xlsx.WorkbookWriter(options);

  const outSheet = outXls.addWorksheet('BIN + PM');
  groupedRows.forEach(rows => {
    rows.forEach(row => {
      const newRow = outSheet.addRow([row.date.toDate(), ...row.vals])
      newRow.commit()
    })
  })
  outSheet.commit()
  outXls.commit()
}
