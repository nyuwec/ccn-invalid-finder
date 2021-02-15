import * as avgOneFile from './avg'
import * as avgFolder from './avg_folder'
import * as listInvalidFiles from './finder'
import * as sliceFile from './slice_file'

const AvailCommands = ['avg', 'avg-folder', 'finder', 'slice-file']

const command = process.argv[2]

if (command == null || AvailCommands.includes(command) == false) {
  console.error("ERR: Please define the command you want to execute. Available commands:")
  AvailCommands.forEach(cmd => console.error(`\t- ${cmd}`))
  console.error(`EXAMPLE: ccn-toolset avg /path/to/file`)
  process.exit(9)
}

const argv = [process.argv[0], process.argv[1], ...process.argv.slice(3)]

console.error(command, argv)

// switch (command) {
//   case 'avg':
//     avgOneFile.default(argv)
//     break
//   case 'avg-folder':
//     avgFolder.default(argv)
//     break
//   case 'finder':
//     listInvalidFiles.default(argv)
//     break
//   case 'slice-file':
//     sliceFile.default(argv)
//     break
//   default:
//     console.error("Command could not be run... weird.")
// }
