/* Read excel file */
const excel = require('exceljs'); // Import class from excelJS package/library

const workbook = new excel.Workbook(); // Creating new object of exceljs class & invoking its method | It hold collection of sheets
workbook.xlsx.readFile("resources/excelTest.xlsx").then(function () {  // Handel the promise 
    const worksheet = workbook.getWorksheet('Sheet1'); // It hold details of entire worksheet
    /* Traversing rows & columns of excel worksheet */
    worksheet.eachRow((row, rowNumber) => { // Outer foreach loop
        row.eachCell((cell, colNumber) => { // Inner foreach loop
            console.log(cell.value);
        })
    })
})

// Another way - by using await
async function excelTest() {
    let coordinate = { row: -1, column: -1 };

    const workbook = new excel.Workbook();
    await workbook.xlsx.readFile("resources/excelTest.xlsx"); // readFile & writeFile are asyncronous operations - for js to wait to complete operation
    const worksheet = workbook.getWorksheet('Sheet1');

    /* Search an element in file */
    worksheet.eachRow((row, rowNumber) => { // Outer loop
        row.eachCell((cell, colNumber) => { // Inner loop
            if (cell.value === "Yellow") {
                console.log(rowNumber + ":" + colNumber); // Print cell coordinates of searched item cell
                coordinate.row = rowNumber;
                coordinate.column = colNumber;
            }
        })
    })

    /* Write in excel files */
    const cell = worksheet.getCell(coordinate.row, coordinate.column);
    cell.value = "Android";
    await workbook.xlsx.writeFile("resources/excelTest.xlsx"); // To save the excel file
}
excelTest();
