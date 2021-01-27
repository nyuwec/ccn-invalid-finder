import * as fs from 'fs'
import * as path from 'path'
import * as parse from 'csv-parse/lib/sync'

export default function listInvalidFiles(dirname: string): {invalidFiles: string[], allFiles: string[]} {
  let allFiles: fs.Dirent[] = fs.readdirSync(dirname, {
    encoding: 'utf8',
    withFileTypes: true
  }).sort()
  let invalidFiles = allFiles//.slice(0, 20)
    .filter((dirEntity) => {
      if (dirEntity.isFile() && path.extname(dirEntity.name) == '.csv') {
        let absPath = dirname + '/' + dirEntity.name
        let data = fs.readFileSync(absPath, 'utf8')
        let records = <Array<Array<string>>>parse(data, {fromLine: 7})
        let invalidSampleFlow = records.some(el => parseFloat(el[17]) < 40.0)
        let invalidAlarmCode = records.some(el => parseFloat(el[47]) > 0.0)

        return (invalidSampleFlow || invalidAlarmCode)
      } else {
        return false
      }
    })

  return {
    invalidFiles: invalidFiles.map((_) => _.name),
    allFiles: allFiles.map((_) => _.name)
  }
}
