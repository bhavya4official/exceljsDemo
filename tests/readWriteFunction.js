/* Read excel file */
const excel = require('exceljs'); // Import class from excelJS package/library

async function writeExcel(searchText, replaceText, change, filePath) {

    const workbook = new excel.Workbook();
    await workbook.xlsx.readFile(filePath); // readFile & writeFile are asyncronous operations - for js to wait to complete operation
    const worksheet = workbook.getWorksheet('Sheet1');

    const coordinate = await readExcel(worksheet, searchText); // Getting corrdinate of searchText cell

    /* Write in excel files */
    const cell = worksheet.getCell(coordinate.row, coordinate.column + change.colChange);
    cell.value = replaceText;
    await workbook.xlsx.writeFile(filePath); // To save the excel file
}

/* Search an element in file */
async function readExcel(worksheet, searchText) {
    let coordinate = { row: -1, column: -1 };

    worksheet.eachRow((row, rowNumber) => { // Outer loop
        row.eachCell((cell, colNumber) => { // Inner loop
            if (cell.value === searchText) {
                console.log(rowNumber + ":" + colNumber); // Printing cell coordinates of searched item
                coordinate.row = rowNumber;
                coordinate.column = colNumber;
            }
        })
    })
    return coordinate;
}

writeExcel("Yellow", Math.floor(Math.random() * 15), { rowChange: 0, colChange: 0 }, "resources/excelTest.xlsx")

// Update Mango price to 350
writeExcel("Mango", 350, { rowChange: 0, colChange: 2 }, "resources/excelTest.xlsx");

