const Crawler = require('crawler')
const xlsx = require('xlsx')
const path = require('path')


// 从xlsx文件第一行中得到CAS，编码成PUG的地址，得到CID写入xlsx中
const workbook = xlsx.readFile(path.join(__dirname, '常见化学品补充3.xlsx'))
let worksheet = workbook.Sheets[workbook.SheetNames[0]]
let CASlists = []
const startRow = 2
const endRow = 8948
let urls = []
let rowth = []
for (let i = startRow; i <= endRow; i++) {
    CASlists.push(worksheet['A' + i]['v'])
    rowth.push(i)
}
for (let i = 0; i < CASlists.length; i++) {
    urls.push({
        url: 'https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/' + CASlists[i] + '/cids/JSON',
        CAS: CASlists[i]
    })
}
console.log(urls);

var n = startRow
const c = new Crawler({
    maxConnections: 1,
    rateLimit: 1500,
    callback: function (error, res, done) {
        if (error) {
            console.log(error)
        } else if (JSON.parse(res.body)['Fault'] == undefined) {
            console.log(res.options.CAS);
            let CID = JSON.parse(res.body)['IdentifierList']['CID'];
            let CAS = res.options.CAS;
            let workbook = xlsx.readFile(path.join(__dirname, '常见化学品补充3结果.xlsx'));
            let worksheet = workbook.Sheets[workbook.SheetNames[0]];
            xlsx.utils.sheet_add_aoa(worksheet, [[CAS, CID]], {
                origin: 'A' + n
            })
            xlsx.writeFile(workbook, path.join(__dirname, '常见化学品补充3结果.xlsx'));
        }
        n = n + 1;
        done();
    }
})
c.queue(urls)



// const workbook = xlsx.readFile(path.join(__dirname, 'a.xlsx'))
// var worksheet = workbook.Sheets[workbook.SheetNames[0]]
// var urls = []
// //  分子重量
// var chemistry = {
//     Name: "",
//     MolecularFomula: "",  //
//     MolecularWeight: "",  //
//     SMILES: "",  //
//     BoilingPoint: "",  //
//     MeltingPoint: "",  //
//     Solubility: "",  //
//     Density: "",  //
//     VaporPressure: "",  //
//     LogP: "",  //
//     pH: "",  //
//     pKa: "",  //
//     rowNumber: 0,
// }

// var endRow = 2871
// var startRow = 2
// for (var i = startRow; i <= endRow; i++) {
//     urls.push('https://pubchem.ncbi.nlm.nih.gov/rest/pug_view/data/compound/' + worksheet['B' + i]['v'] + '/JSON/')
// }

// function NamesAndIdentifiers(inputData, outputData) {
//     for (var j = 0; j < inputData['Section'].length; j++) {
//         if (inputData['Section'][j]['TOCHeading'] == 'Molecular Formula') {
//             outputData['MolecularFomula'] = inputData['Section'][j]['Information'][0]['Value']['StringWithMarkup'][0]['String']
//         }
//         if (inputData['Section'][j]['TOCHeading'] == 'Computed Descriptors') {
//             for (var m = 0; m < inputData['Section'][j]['Section'].length; m++) {
//                 if (inputData['Section'][j]['Section'][m]['TOCHeading'] == 'Canonical SMILES') {
//                     outputData['SMILES'] = inputData['Section'][j]['Section'][m]['Information'][0]['Value']['StringWithMarkup'][0]['String']
//                 }
//             }
//         }
//         if (inputData['Section'][j]['TOCHeading'] == 'Other Identifiers') {
//             for (var m = 0; m < inputData['Section'][j]['Section'].length; m++) {
//                 if (inputData['Section'][j]['Section'][m]['TOCHeading'] == 'CAS') {
//                     outputData['CAS'] = inputData['Section'][j]['Section'][m]['Information'][0]['Value']['StringWithMarkup'][0]['String']
//                 }
//             }
//         }
//     }
// }
// function ChemicalPorperty(inputData, outputData) {
//     for (var j = 0; j < inputData['Section'].length; j++) {
//         if (inputData['Section'][j]['Section'] != undefined) {
//             for (var n = 0; n < inputData['Section'][j]['Section'].length; n++) {
//                 if (inputData['Section'][j]['Section'][n]['TOCHeading'] == 'Molecular Weight') {
//                     if (inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'] == undefined) {
//                         outputData['MolecularWeight'] = '' + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Number'][0] + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Unit']
//                     } else {
//                         outputData['MolecularWeight'] = inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'][0]['String']
//                     }
//                 }
//                 if (inputData['Section'][j]['Section'][n]['TOCHeading'] == 'Boiling Point') {
//                     if (inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'] == undefined) {
//                         outputData['BoilingPoint'] = '' + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Number'][0] + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Unit']
//                     } else {
//                         outputData['BoilingPoint'] = inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'][0]['String']
//                     }
//                 }
//                 if (inputData['Section'][j]['Section'][n]['TOCHeading'] == 'Melting Point') {
//                     if (inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'] == undefined) {
//                         outputData['MeltingPoint'] = '' + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Number'][0] + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Unit']
//                     } else {
//                         outputData['MeltingPoint'] = inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'][0]['String']
//                     }
//                 }
//                 if (inputData['Section'][j]['Section'][n]['TOCHeading'] == 'Solubility') {
//                     if (inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'] == undefined) {
//                         outputData['Solubility'] = '' + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Number'][0] + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Unit']
//                     } else {
//                         outputData['Solubility'] = inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'][0]['String']
//                     }
//                 }
//                 if (inputData['Section'][j]['Section'][n]['TOCHeading'] == 'Density') {
//                     if (inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'] == undefined) {
//                         outputData['Density'] = '' + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Number'][0] + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Unit']
//                     } else {
//                         outputData['Density'] = inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'][0]['String']
//                     }
//                 }
//                 if (inputData['Section'][j]['Section'][n]['TOCHeading'] == 'Vapor Pressure') {
//                     if (inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'] == undefined) {
//                         outputData['VaporPressure'] = '' + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Number'][0] + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Unit']
//                     } else {
//                         outputData['VaporPressure'] = inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'][0]['String']
//                     }
//                 }
//                 if (inputData['Section'][j]['Section'][n]['TOCHeading'] == 'LogP') {
//                     for (var m = 0; m < inputData['Section'][j]['Section'][n]['Information'].length; m++) {
//                         if (inputData['Section'][j]['Section'][n]['Information'][m]['Value']['StringWithMarkup'] == undefined) {
//                             outputData['LogP'] = outputData['LogP'] + inputData['Section'][j]['Section'][n]['Information'][m]['Value']['Number'][0] + "|"
//                         } else {
//                             outputData['LogP'] = outputData['LogP'] + inputData['Section'][j]['Section'][n]['Information'][m]['Value']['StringWithMarkup'][0]['String'] + "|"
//                         }
//                     }
//                 }
//                 if (inputData['Section'][j]['Section'][n]['TOCHeading'] == 'pH') {
//                     if (inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'] == undefined) {
//                         outputData['pH'] = '' + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Number'][0] + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Unit']
//                     } else {
//                         outputData['pH'] = inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'][0]['String']
//                     }
//                 }
//                 if (inputData['Section'][j]['Section'][n]['TOCHeading'] == 'pKa') {
//                     if (inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'] == undefined) {
//                         outputData['pKa'] = '' + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Number'][0] + inputData['Section'][j]['Section'][n]['Information'][0]['Value']['Unit']
//                     } else {
//                         outputData['pKa'] = inputData['Section'][j]['Section'][n]['Information'][0]['Value']['StringWithMarkup'][0]['String']
//                     }
//                 }
//             }
//         }
//     }
// }
// function refreshChemistry(chemistry) {
//     addToExcel(chemistry, chemistry['rowNumber'])
//     chemistry['CAS'] = ''
//     chemistry['CID'] = ''
//     chemistry['Name'] = ''
//     chemistry['MolecularFomula'] = ''
//     chemistry['MolecularWeight'] = ''
//     chemistry['SMILES'] = ''
//     chemistry['BoilingPoint'] = ''
//     chemistry['MeltingPoint'] = ''
//     chemistry['Solubility'] = ''
//     chemistry['Density'] = ''
//     chemistry['VaporPressure'] = ''
//     chemistry['LogP'] = ''
//     chemistry['pH'] = ''
//     chemistry['pKa'] = ''
//     chemistry['rowNumber'] = 0

// }

// function addToExcel(Chemistry, startRow) {
//     var lists = []
//     lists.push(Chemistry['CAS'])
//     lists.push(Chemistry['CID'])
//     lists.push(Chemistry['Name'])
//     lists.push(Chemistry['SMILES'])
//     lists.push(Chemistry['MolecularFomula'])
//     lists.push(Chemistry['MolecularWeight'])
//     lists.push(Chemistry['BoilingPoint'])
//     lists.push(Chemistry['MeltingPoint'])
//     lists.push(Chemistry['Solubility'])
//     lists.push(Chemistry['Density'])
//     lists.push(Chemistry['VaporPressure'])
//     lists.push(Chemistry['LogP'])
//     lists.push(Chemistry['pH'])
//     lists.push(Chemistry['pKa'])
//     var inputbook = xlsx.readFile(path.join(__dirname, 'b.xlsx'))
//     var inputsheet = inputbook.Sheets[inputbook.SheetNames[0]]
//     xlsx.utils.sheet_add_aoa(inputsheet, [lists], {
//         origin: 'A' + startRow
//     })
//     console.log(lists[0]);
//     xlsx.writeFile(inputbook, path.join(__dirname, 'b.xlsx'))
// }

// var c = new Crawler({
//     rateLimit: 2000,
//     maxConnections: 1,
//     callback: function (error, res, done) {
//         if (error) {
//             console.log(error)
//         } else {
//             var data = JSON.parse(res.body)['Record']['Section']
//             for (var t = 0; t < data.length; t++) {
//                 if (data[t]['TOCHeading'] == 'Chemical and Physical Properties') {
//                     ChemicalPorperty(data[t], chemistry)
//                 }
//                 if (data[t]['TOCHeading'] == 'Names and Identifiers') {
//                     NamesAndIdentifiers(data[t], chemistry)
//                 }
//                 chemistry.Name = JSON.parse(res.body)['Record']['RecordTitle']
//             }
//             chemistry['rowNumber'] = startRow
//             chemistry['CID'] = JSON.parse(res.body)['Record']['RecordNumber']
//             refreshChemistry(chemistry)
//         }

//         startRow = startRow + 1
//         done()
//     }
// })

// c.queue(urls)
