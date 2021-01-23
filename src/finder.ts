import * as fs from 'fs'
import * as parse from 'csv-parse/lib/sync'

export default function listInvalidFiles(dirname: string): {invalidFiles: string[], allFiles: string[]} {
  let allFiles = fs.readdirSync(dirname).sort()
  let invalidFiles = allFiles//.slice(0, 20)
    .filter((filename) => {
      let path = dirname + '/' + filename
      let data = fs.readFileSync(path, 'utf8')
      let records = <Array<Array<string>>>parse(data, {fromLine: 7})
      let invalidSampleFlow = records.some(el => parseFloat(el[17]) < 40.0)
      let invalidAlarmCode = records.some(el => parseFloat(el[47]) > 0.0)

      return (invalidSampleFlow || invalidAlarmCode)
    })

  return {
    invalidFiles: invalidFiles,
    allFiles: allFiles
  }
}
