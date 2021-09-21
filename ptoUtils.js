const xlsx = require('xlsx')

/**
* Read an excel book and convert a single sheet to a json object
*
* @param {String} excelFileName - workbook to read
* @param {String} sheetName - sheet to convert to json
* @return {Object} Json object containing sheet of data from workbook
*/
function excelToJson(excelFileName,sheetName) {
	// Read the excel workbook
	const book = xlsx.readFile(excelFileName)
	// Grab the sheet of interest
	const sheet = book.Sheets[sheetName]
	// Convert sheet to usable json.
	return xlsx.utils.sheet_to_json(sheet)
}
function getAllTables(inputs) {
	console.log('Get All Tables')
	var tables = {}
	inputs.forEach(input => { 
    // No file to read means put an empty table on the list
    if(typeof input.file === 'undefined') { 
      tables[input.attribute] = [] 
    } else{
      // Read the table from an excel spreadsheet
      tables[input.attribute] = excelToJson(input.file,input.sheet) 
      tables[input.attribute].forEach(row => {
        if (input.newfields) {
          input.newfields.forEach(newfield => { row[newfield] = null })
        }
        var key =
          ( row[input.column] ) 
            ? (row[input.column] || '').toUpperCase().replace('.USPTO.GOV','').trim()
            : '';
        if (input.specialConditioning) { 
          key = eval(input.specialConditioning) 
        }
        row._key = key
      })
    } // endif (no input file)
    console.log('    %s conditioned',(typeof input.file === 'undefined') ? input.attribute : input.file)
	})
	return tables
}


function writeSheetToBook(sheets) {
  var book = xlsx.utils.book_new();
  var bookName = './data/test/temp' + Date.now() + '.xlsx'
  var timerPrompt = 'Creating Book [' + bookName + ']'
  console.time(timerPrompt)

  sheets.forEach(sheetRow => {
    console.log('Converting Sheet [%s]',sheetRow.name)
    var sheet = xlsx.utils.json_to_sheet(sheetRow.data);
    xlsx.utils.book_append_sheet(book, sheet, sheetRow.name)
  })
	xlsx.writeFile(book, bookName);

  console.timeEnd(timerPrompt)
}
module.exports.getAllTables = getAllTables
module.exports.writeSheetToBook = writeSheetToBook
