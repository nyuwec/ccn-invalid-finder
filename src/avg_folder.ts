import * as fs from 'fs'
import * as path from 'path'
import * as Excel from 'exceljs'
import * as moment from 'moment'
import { DateRow, GroupedRows, calculateAvg, writeAvgs, extractNumber, toMoment, Cols } from './models/avg'

export default function avgFolder(argv: string[]) {
  const paramFolderName = argv[2]

  if (paramFolderName == null) {
    console.error("ERR: Please define all the params:")
    console.error("\t- path to folder")
    console.error(`EXAMPLE: ./${process.env.npm_package_name} avg-folder /path/to/`)
    process.exit(0)
  }

  const folderName: string = (paramFolderName + '/').replace(/\/\/$/, '/')

  const START_ROW = 9
  const BIN_FIRST_COL = 47
  const BIN_LAST_COL = 71
  const PM_FIRST_COL = 41
  const PM_LAST_COL = 46

  loadDataFromFolder(folderName)
    .then(resultedGRows => {
      const lastFolderSegment = folderName.split('/').slice(-2, -1)
      console.log('Merge GroupedRows...')
      const groupedRows: GroupedRows = new GroupedRows()
      resultedGRows.forEach(gr => {
        groupedRows.append(gr)
      })
      console.log('Writing out GroupedRows...')
      writeGroupedRows(groupedRows, folderName + `/${lastFolderSegment}_full.xlsx`)

      console.log('Writing out 10Min avgs...')
      const avgResult = calculateAvg(groupedRows)
      writeAvgs(avgResult, folderName + `/${lastFolderSegment}_full_10MinAVG.xlsx`)
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

    let result = Array<GroupedRows>()
    for await (const r of readFiles(all)) {
      result.push(r)
    }

    return result
  }

  async function* readFiles(files: string[]) {
    for(const file of files) {
      yield readFile(file);
    }
  };

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
    const date: moment.Moment = toMoment(rawRow.getCell(Cols.DATE))

    const vals: number[] = Array<number>()
    for (let i=BIN_FIRST_COL;i<=BIN_LAST_COL;i++) {
      vals.push(extractNumber(rawRow.getCell(i)))
    }
    for (let i=PM_FIRST_COL;i<=PM_LAST_COL;i++) {
      vals.push(extractNumber(rawRow.getCell(i)))
    }
    return {
      date: date,
      vals: vals
    }
  }

  function writeGroupedRows(groupedRows: GroupedRows, fileName: string) {
    const filePath = path.parse(fileName)
    const maxLinesPerFile = 300000
    let fileCount = 0
    let currLine = 0

    let outXls = genWriter(filePath, fileCount)
    let outSheet = genSheet(outXls)
    groupedRows.forEach(rows => {
      rows.forEach(row => {
        const newRow = outSheet.addRow([row.date.toDate(), ...row.vals])
        newRow.commit()
        currLine++
        if (currLine >= maxLinesPerFile) {
          outSheet.commit()
          outXls.commit()

          fileCount++
          currLine = 0
          outXls = genWriter(filePath, fileCount)
          outSheet = genSheet(outXls)
        }
      })
    })
    outSheet.commit()
    outXls.commit()

    function genWriter(filePath: path.ParsedPath, fileCount: number) {
      const options = {
        filename: `${filePath.dir}/${filePath.name}_${fileCount.toString().padStart(3, '0')}${filePath.ext}`,
        useStyles: false,
        useSharedStrings: false
      }
      const outXls = new Excel.stream.xlsx.WorkbookWriter(options)
      return outXls
    }
    function genSheet(outXls: Excel.stream.xlsx.WorkbookWriter) {
      return outXls.addWorksheet('BIN + PM')
    }
  }
}
