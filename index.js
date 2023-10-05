//excel4node
//xlsx
//json2xls

/**
 *  @file converts json data into excel file
 *  @author Ronak Jagani
 *  @see {@link https://github.com/kevit-ronak-jagani/jsonToExcel.git|GitHub Repo}
 */

// Import the library
const xlsx = require('xlsx');

/**
 * Read the JSON data from a file.
 * @type {Array} - The JSON data representing customer information.
 * @see {@link ./data.json|Created json Data}
 */

const jsonData = require('./data.json');

// Extract required useful data fields
const extractedData = jsonData.map((customer) => {
    /**
     * @property {string} first - The first name of the customer.
     * @property {string} last - The last name of the customer.
     * @property {string} email - The email address of the customer.
     * @property {number} age - The age of the customer.
     */

    // Destructuring customer object to get name, email, and dateOfBirth fields
    const { name, email, dateOfBirth } = customer;
    const { first, last } = name;

    // Calculate age based on date of birth
    const today = new Date();
    const birthDate = new Date(dateOfBirth);
    const age = today.getFullYear() - birthDate.getFullYear();

    // Return an array with the required fields
    return [first, last, email, age];
});

// Create a workbook and add a worksheet
const wb = xlsx.utils.book_new();
const ws = xlsx.utils.aoa_to_sheet([['First Name', 'Last Name', 'Email', 'Age'], ...extractedData]);
xlsx.utils.book_append_sheet(wb, ws, 'Customers');

// Save the Excel file
xlsx.writeFile(wb, 'customer_data.xlsx');

console.log('Excel file created');


