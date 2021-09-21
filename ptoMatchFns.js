function ExactCyberToGEARSServerMatch(tables) {
	var ecmoTable = tables['ecmo']
	var gearsTable = tables['gears']

	console.time('ExactCyberToGEARSServerMatch')
	ecmoTable.forEach(ecmoRow => {
		gearsTable.forEach(gearsRow => {
			if(ecmoRow._key == gearsRow._key) {
				ecmoRow['GEARS EXACT MATCH'] = gearsRow._key
			}
		})
	})
	console.log('ExactCyberToGEARSServerMath=%d matched',ecmoTable.filter(row=>{return row['GEARS EXACT MATCH'] != null}).length)
	console.timeEnd('ExactCyberToGEARSServerMatch')
}

function IdentifyType(tables) {
	var ecmoTable = tables['ecmo']
	console.time('IdentifyType')
	ecmoTable.forEach(ecmoRow => {
		if( ecmoRow.Relay.match(/dev-/i)) { ecmoRow.Label = 'dev' }
		else if( ecmoRow.OS.match(/win10/i)) { ecmoRow.Label = 'dev'}
		else if(ecmoRow['GEARS EXACT MATCH']) { ecmoRow.Label = 'Exact Match'}
	})
	console.log('IdentifyType=%d matched',ecmoTable.filter(row=>{return row.Label != null}).length)
	console.timeEnd('IdentifyType')
}

function MoveCyberNonProd(tables) {
	console.time('MoveCyberNonProd')
	var ecmoTable = tables['ecmo']
	console.log('MoveCyberNonProd=%d ON ENTRY',ecmoTable.length)
	tables['dev'] = ecmoTable.filter(ecmoRow => { return ecmoRow.Label === 'dev' })
	tables['exactmatch'] = ecmoTable.filter(ecmoRow => { return ecmoRow.Label === 'Exact Match' })
	tables['ecmo'] = ecmoTable.filter(ecmoRow => { return ecmoRow.Label === null })
	console.log('MoveCyberNonProd-DEV=%d matched',tables['dev'].length)
	console.log('MoveCyberNonProd-EXACTMATCH=%d matched',tables['exactmatch'].length)
	console.log('MoveCyberNonProd-ECMO=%d AFTER REMOVE',ecmoTable.length)
	console.timeEnd('MoveCyberNonProd')
}

function MatchOnSDAP(tables) {
	var ecmoTable = tables['ecmo']
	var sdapTable = tables['sdap']
	console.time('MatchOnSDAP')
	ecmoTable.forEach(ecmoRow => {
		sdapTable.forEach(sdapRow => {
			if(ecmoRow._key == sdapRow._key) {
				if (sdapRow['ADDED: Component Name']) {
					ecmoRow['SDAP Exact Match'] = 
						(sdapRow['ADDED: Component Name'] || '')
							.toUpperCase()
							.trim()
				}
				
			}
		})
	})
	console.log('MatchOnSDAP=%d matched',ecmoTable.filter(row=>{return row['SDAP Exact Match'] != null}).length)
	console.timeEnd('MatchOnSDAP')
}

function CheckManual(tables) {
	var ecmoTable = tables['ecmo']
	var manualTable = tables['manual']
	console.time('CheckManual')
	ecmoTable.forEach(ecmoRow => {
		manualTable.forEach(manualRow => {
			if(ecmoRow._key == manualRow._key) {
				if ( manualRow['Component'] ) {
					ecmoRow['Component Manual'] = 
						(manualRow['Component'] || '')
							.toUpperCase()
							.trim()
				}
				
			}
		})
	})
	console.log('CheckManual=%d matched',ecmoTable.filter(row=>{return row['Component Manual'] != null}).length)
	console.timeEnd('CheckManual')
}

function DiamondCheck(tables) {
	var ecmoTable = tables['ecmo']
	var diamondTable = tables['diamond']
	console.time('DiamondCheck')
	ecmoTable.forEach(ecmoRow => {
		diamondTable.forEach(diamondRow => {
			if(ecmoRow._key == diamondRow._key) {
				if ( diamondRow['GEARS Component'] ) {
					ecmoRow['Diamond Lookup'] =
						(diamondRow['GEARS Component'] || '')
							.toUpperCase()
							.trim()
				}
				
			}
		})
	})
	console.log('DiamondCheck=%d matched',ecmoTable.filter(row=>{return row['Diamond Lookup'] != null}).length)
	console.timeEnd('DiamondCheck')
}

module.exports.ExactCyberToGEARSServerMatch = ExactCyberToGEARSServerMatch
module.exports.IdentifyType = IdentifyType
module.exports.MoveCyberNonProd = MoveCyberNonProd
module.exports.MatchOnSDAP = MatchOnSDAP
module.exports.CheckManual = CheckManual
module.exports.DiamondCheck = DiamondCheck