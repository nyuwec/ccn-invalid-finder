import * as fs from 'fs'
import * as path from 'path'
import * as parse from 'csv-parse/lib/sync'

export default function finder(argv: string[]) {
  let dirname = argv[2]

  if (dirname == null) {
    console.error("ERR: Please define all the params:")
    console.error("\t- path to folder")
    console.error(`EXAMPLE: ./${process.env.npm_package_name} finder /path/to/`)
    process.exit(0)
  }

  try {
    console.log(`Listing files in: "${dirname}"`)
    let result = listInvalidFiles(dirname)

    console.log(`Invalid files:`)
    result.invalidFiles.forEach(fn => console.log(fn))
    console.log(`invalids: ${result.invalidFiles.length}`)
    console.log(`     all: ${result.allFiles.length}`)
  } catch (error) {
    console.log(`ERR:`, error)
  }
}


function listInvalidFiles(dirname: string): {invalidFiles: string[], allFiles: string[]} {
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
