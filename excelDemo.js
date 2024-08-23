/* Read excel file */
const excel = require('exceljs'); // Import class from excelJS package/library

const workbook = new excel.Workbook(); // Creating new object of exceljs class & invoking its method | It hold collection of sheets
workbook.xlsx.readFile("resources/excelTest.xlsx").then(function () {  // Handel the promise 
    const worksheet = workbook.getWorksheet('Sheet1'); // It hold details of entire worksheet
    /* Traversing rows & columns of excel worksheet */
    worksheet.eachRow((row, rowNumber) => { // Outer loop
        row.eachCell((cell, colNumber) => { // Inner loop
            console.log(cell.value);
        })
    })
})

// Another way - by using await
async function excelTest() {
    const workbook = new excel.Workbook();
    await workbook.xlsx.readFile("resources/excelTest.xlsx");  // Wait until this operation is completed then proceed to next
    const worksheet = workbook.getWorksheet('Sheet1'); // It hold details of entire worksheet

    worksheet.eachRow((row, rowNumber) => { // Outer loop
        row.eachCell((cell, colNumber) => { // Inner loop
            console.log(cell.value);
        })
    })
}
excelTest();
