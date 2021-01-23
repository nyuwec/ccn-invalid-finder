import listInvalidFiles from './finder'

let dirname = process.argv[2] /* what the user enters as first argument */

if (dirname == null) {
  console.error("ERR: Please pass the directory name parameter.")
  process.exit(9)
}

try {
  let result = listInvalidFiles(dirname)

  console.log(`Invalid files:`)
  result.invalidFiles.forEach(fn => console.log(fn))
  console.log(`invalids: ${result.invalidFiles.length}`)
  console.log(`     all: ${result.allFiles.length}`)
} catch (error) {
  console.log(`ERR:`, error)
}
