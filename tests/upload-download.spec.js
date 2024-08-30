import { test, expect } from '@playwright/test';

const excel = require('exceljs'); // Import class from excelJS package/library

/* Read excel file */
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
            else {
                console.log("Search text not found in excel file.");
            }
        })
    })
    return coordinate;
}

// writeExcel("Yellow", Math.floor(Math.random() * 15), { rowChange: 0, colChange: 0 }, "resources/excelTest.xlsx")
// Update Mango price to 350
// writeExcel("Mango", 350, { rowChange: 0, colChange: 2 }, "resources/excelTest.xlsx");

test("Upload download excel validation", async function ({ page }) {
    const textSearch = 'Mango';
    const updatedValue = '240';

    /* Download excel file */
    await page.goto("https://rahulshettyacademy.com/upload-download-test/index.html");
    const downloadPromise = page.waitForEvent('download'); // Call that event to happen - keep an eye
    await page.getByRole('button', { name: 'Download' }).click();
    await downloadPromise; // Wait for download event to complete - wait until download promis to resolved

    await writeExcel(textSearch, updatedValue, { rowChange: 0, colChange: 2 }, "C:/Users/Bhavya Singh/Downloads/Chrome/download.xlsx");

    /* Upload excel file on site */
    await page.locator('#fileinput').click();
    await page.locator('#fileinput').setInputFiles("C:/Users/Bhavya Singh/Downloads/Chrome/download.xlsx");

    /* Expect Mango price on site after uploading file */
    // await expect(page.locator('#row-0 #cell-4-undefined div')).toHaveText(240);
    const textLocator = page.getByText(textSearch);
    const desiredRow = page.getByRole('row').filter({ has: textLocator });
    await expect(desiredRow.locator('#cell-4-undefined')).toContainText(updatedValue);

});