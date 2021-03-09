import avg10mFile from './avg'
import avgFile1min from './avg1min'
import avg10mFolder from './avg_folder'
import finder from './finder'
import sliceFile from './slice_file'

const AvailCommands = ['avg10m', 'avg1m', 'avg10m-folder', 'finder', 'slice-file']

const command = process.argv[2]

if (command == null || AvailCommands.includes(command) == false) {
  console.error("ERR: Please define the command you want to execute. Available commands:")
  AvailCommands.forEach(cmd => console.error(`\t- ${cmd}`))
  console.error(`EXAMPLE: ./${process.env.npm_package_name} avg /path/to/file`)
  process.exit(9)
}

const argv = [process.argv[0], process.argv[1], ...process.argv.slice(3)]

console.log(`-=- Running ${command.toUpperCase()} -=-`)

switch (command) {
  case 'avg10m':
    avg10mFile(argv)
    break
  case 'avg1m':
    avgFile1min(argv)
    break
  case 'avg10m-folder':
    avg10mFolder(argv)
    break
  case 'finder':
    finder(argv)
    break
  case 'slice-file':
    sliceFile(argv)
    break
  default:
    console.error("Command could not be run... weird.")
}
