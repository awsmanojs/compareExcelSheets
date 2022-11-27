const ExcelJS = require('exceljs')
const XLSX = require('xlsx')
const xlsxFile = require('read-excel-file/node')

//read BU_ID file for the list of CX_BU_ID what user provides.
const workbook = XLSX.readFile("BU_ID.xlsx")

//const worksheet = workbook.Sheets[workbook.SheetNames[0]]

// get the first Sheet data into worksheet variable
const worksheet = workbook.Sheets["Sheet1"]

// store the data into json variable
const arrList = XLSX.utils.sheet_to_json(worksheet)
var IDList = []

// loop through the json variable and store the data into IDList array
for (i = 0; i < arrList.length; i++) {
    IDList.push(arrList[i].CX_BU_ID)
}

console.log(IDList)

xlsxFile('customer_data.xlsx').then((rows) => {
    // `rows` is an array of rows
    // each row being an array of cells.
    console.log(rows)

    // copy the header from second file which has multiple columns
    var twoHeaderArray = rows[0]

    // get the length of the header array
    var twoHeaderArrayLength = twoHeaderArray.length
    console.log(twoHeaderArray)
    console.log(twoHeaderArrayLength)

    // get the index of the column which has CX_BU_ID
    for (j=0; j<twoHeaderArrayLength; j++) {
        if (twoHeaderArray[j] === "CX_BU_ID") {
            var CX_BU_ID_index = j
            console.log("CX_BU_ID is found Index is: " + CX_BU_ID_index)
        }
    }

    // create a variable to store the data which has CX_BU_ID
    var customerList = []

    // loop through the rows and store the data which has CX_BU_ID from second file.
    for (i = 1; i < rows.length; i++) {
        customerList.push(rows[i][CX_BU_ID_index])
    }
    console.log(customerList)

    // create a variable to store data which has CX_BU_ID from first file and second file
    var result = customerList.filter(function (item) {
        return IDList.includes(item);
    });
    console.log(result)

    // create new workbook
    var workbook = new ExcelJS.Workbook();

    // create new worksheet
    var worksheet = workbook.addWorksheet('Sheet1');

    // add header row
    worksheet.columns = [
        { header: 'CX_BU_ID', key: 'CX_BU_ID', width: 30 },
    ];

    // add rows in the new result worksheet
    for (i = 0; i < result.length; i++) {
        worksheet.addRow({ CX_BU_ID: result[i] });
    }

    // save the workbook
    workbook.xlsx.writeFile('result.xlsx')
        .then(function () {
            console.log('file saved!');
        });
})