const ptoUtils = require('./ptoUtils')
const ptoFindServers = require('./FindServersNSLookup')
const ptoFindComponents = require('./FindComponentinPML')
const ptoMatch = require('./ptoMatchFns')
const ptoServerSw = require('./ServerSWUpdates')

String.prototype.left = function(n) {
    return this.substring(0, n);
}
String.prototype.removeNum = function() {
	return this.replace(/\d+/g, '')
}
String.prototype.trimRight = function(n) {
	return this.replace(new RegExp("["+n+"]+$"), '')
}
String.prototype.trimLeft = function(n) {
	return this.replace(new RegExp("^["+n+"]+"), '')
}

console.log('====================================================')
console.log('OCTO SP Site Clean - Initiated')
console.log('----------------------------------------------------')

// List of inputs 
// - file = excel file (will have to change for SP list read)
// - sheet = sheet in excel workbook containing data
// - column = column on worksheet containing computer or server value
// - attribute = what group of data is this (used to reference list)
// - component:name - property containing Component name (optional)
// - component:id - property containing Component id (optional)
const inputs = [
  {file: './data/test/ecmo.xlsx', sheet: 'ECMO', column: 'Computer Name',attribute: 'ecmo',component: { name: '', id: ''}, newfields: ["Label" ,"SDAP EXACT MATCH" ,"GEARS EXACT MATCH" ,"Component Manual" ,"Component from PML" ,"Component from GEARS" ,"LOGIC Match" ,"Diamond lookup" ,"Final Combined" ,"Software Name" ,"Software ID" ,"Database" ,"Server" ,"NSLookup"]}
  ,{file: './data/gears.xlsx', sheet: 'GEARS', column: 'Server',attribute: 'gears',component: { name: '', id: 'ComponentLookup:ComponentID'} }
  ,{file: './data/sdap.xlsx', sheet: 'SDAP', column: 'Hostname',attribute: 'sdap',component: { name: 'ADDED: Component Name', id: 'ADDED: Component ID'}}
  ,{file: './data/manual.xlsx', sheet: 'Manual', column: 'Server Name',attribute: 'manual',component: { name: 'Component', id: ''} }
  ,{file: './data/diamond.xlsx', sheet: 'Diamond', column: 'Hostname',attribute: 'diamond',component: { name: 'GEARS Component', id: 'Component ID'}, specialConditioning: "key.replace(String.fromCharCode(160),' ')"}
  ,{file: './data/pml.xlsx', sheet: 'PML', column: 'ComponentAcronym',attribute: 'pml',component: { name: 'ComponentAcronym', id: 'ID'}}
  ,{attribute: 'comparens', sheet: 'CompareNSLookupServers'}
  ,{attribute: 'cyberlist', sheet: 'CyberListfromCNSLookupS'}
  ,{attribute: 'serversw',sheet: 'ServerSWUpdates'}
  ,{attribute: 'dev',sheet: 'Dev'}
  ,{attribute: 'exactmatch',sheet: 'ExactMatch'}
//  	,{file: './data/server.xlsx', sheet: 'Server', column: 'ServerName',attribute: 'server'}
]

console.time('MAIN PROGRAM')

// Read all excel tables
var tables = ptoUtils.getAllTables(inputs);

// ptoCheckAndMatch.:ExactCyberToGEARSServerMatch()
try 
{
	ptoMatch.ExactCyberToGEARSServerMatch(tables)
	ptoMatch.IdentifyType(tables)
	ptoMatch.MoveCyberNonProd(tables)
	ptoMatch.MatchOnSDAP(tables)
	ptoMatch.CheckManual(tables)
	ptoMatch.DiamondCheck(tables)
	ptoUtils.writeSheetToBook(
		[
			{name: 'Cyber',data:tables['ecmo']}
			,{name: 'GEARS',data:tables['gears']}
			,{name: 'SDAP',data:tables['sdap']}
			,{name: 'Dev',data:tables['dev']}
			,{name: 'ExactMatch',data:tables['exactmatch']}
		])	
	ptoFindComponents.FindComponentinPML(tables)
	ptoServerSw.ServerSWUpdates(tables)
	ptoFindServers.FindServersNSLookup(tables)
} catch(err) {
	console.error(2,err)
} finally {
	console.timeEnd('MAIN PROGRAM')
}


// function upTrim(obj) {
// 	return typeof obj === 'string' ? obj.toUpperCase().trim() : obj
// }
