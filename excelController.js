const xlsx = require('xlsx');   
const fs = require('fs');      
// read -> modify -> jdiin json -> modified object, return this.


// 1 entry point controller
// dalem 1 controller: delete 1 funtion, process 1 function, post 1 function
// user lngsung tau berhasil/ga
// json di tax file
module.exports = {
  deleteRowsController: (req, res) => {
    // Read the Excel file
    const inputWorkbook = xlsx.readFile('sales_report-04-Sep-2023@1693879681.xlsx');
    const sheetName = inputWorkbook.SheetNames[0]; 
    
    let sheet = inputWorkbook.Sheets[sheetName];
    date = getDate(sheet);

    sheet = delete4Rows(sheet);
    sheet = replaceEmptyWithNull(sheet);

    // Create a new workbook with the modified sheet
    const outputWorkbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(outputWorkbook, sheet, sheetName);

    // Save the modified workbook back to the file
    xlsx.writeFile(outputWorkbook, 'modified_sales_report.xlsx');

    // make it to json and send
    const modifiedSheetData = xlsx.utils.sheet_to_json(sheet);
    
    // modify date function to turn the excel format into retailsoft's
    date = modifyDate(date);
    console.log(date)

    /* process the json data into retailsoft version */
    function transformExcelData(originalData) {
        const transformedData = [];

        // Create a new object to hold the transformed data for this entry
        let transformedEntry = {};
          
        // Step 3: Extract values from the original data and format them
        transformedEntry.no = `JV-0001`;
        transformedEntry.dt = date;
        transformedEntry.reviewDate = date;
        transformedEntry.approvedDate = date;
        transformedEntry.type = "1";
        transformedEntry.locCode = "HO";
        transformedEntry.createBy = "gusto";
        transformedEntry.reviewedBy = "gusto";
        transformedEntry.approvedBy = "gusto";
        transformedEntry.remark = `Sales Gusto ${date}`;
        transformedEntry.details = [];
        
        
        // handle the details --> every entry becoming 5 details
        for (const entry of originalData) {
            
            const detailDebit = {
                accountCode: "1101998", 
                debitCredit: 1,
                amount: entry["Rate"], 
                debitAmount: entry["Rate"],
                creditAmount: "0",
                currencyCode: "IDR",
                currencyRate: 1,
                description: `${entry["Guest Name"]} Folio: ${entry["Folio No."]} Room: ${entry["Room No."]} Source : ${entry["Source"]}`,
                locCode: "HO",
                deptCode: "",
                prjCode: ""
            };
              
            // Add the detail object to the 'details' array
            transformedEntry.details.push(detailDebit);

            commission = entry["Rate"] * .192

            const detailCommission = {
                accountCode: "6011023", 
                debitCredit: 2,
                amount: commission, 
                debitAmount: "0",
                creditAmount: commission,
                currencyCode: "IDR",
                currencyRate: 1,
                description: `${entry["Guest Name"]} Folio: ${entry["Folio No."]} Room: ${entry["Room No."]} Source : ${entry["Source"]}`,
                locCode: "HO",
                deptCode: "",
                prjCode: ""
            };
              
            // Add the detail object to the 'details' array
            transformedEntry.details.push(detailCommission);

            const serviceCharge = (entry["Rate"] - commission) * .08264 // assuming svc is ~8% (taken from a few samples)

            const detailService = {
                accountCode: "4011001", 
                debitCredit: 2,
                amount: serviceCharge, 
                debitAmount: "0",
                creditAmount: serviceCharge,
                currencyCode: "IDR",
                currencyRate: 1,
                description: `${entry["Guest Name"]} Folio: ${entry["Folio No."]} Room: ${entry["Room No."]} Source : ${entry["Source"]}`,
                locCode: "HO",
                deptCode: "",
                prjCode: ""
            };
              
            // Add the detail object to the 'details' array
            transformedEntry.details.push(detailService);
            const taxPb = (entry["Rate"] - commission - serviceCharge) * .11

            const detailTax = {
                accountCode: "2500001", 
                debitCredit: 2,
                amount: taxPb, 
                debitAmount: "0",
                creditAmount: taxPb,
                currencyCode: "IDR",
                currencyRate: 1,
                description: `${entry["Guest Name"]} Folio: ${entry["Folio No."]} Room: ${entry["Room No."]} Source : ${entry["Source"]}`,
                locCode: "HO",
                deptCode: "",
                prjCode: ""
            };
              
            // Add the detail object to the 'details' array
            transformedEntry.details.push(detailTax);

            const roomRevenue = entry["Rate"] - commission - serviceCharge - taxPb

            const detailRoomRevenue = {
                accountCode: "2103006", // still unclear
                debitCredit: 2,
                amount: roomRevenue, 
                debitAmount: "0",
                creditAmount: roomRevenue,
                currencyCode: "IDR",
                currencyRate: 1,
                description: `${entry["Guest Name"]} Folio: ${entry["Folio No."]} Room: ${entry["Room No."]} Source : ${entry["Source"]}`,
                locCode: "HO",
                deptCode: "",
                prjCode: ""
            };
              
            // Add the detail object to the 'details' array
            transformedEntry.details.push(detailRoomRevenue);

        }    
        // Add the transformed entry to the 'transformedData' array
        transformedData.push(transformedEntry);

        return transformedData;
    }
      
    // Call the function to transform the data
    const transformedResult = transformExcelData(modifiedSheetData);
      
     // Print the transformed data
    // console.log(JSON.stringify(transformedResult, null, 2));
    // res.status(200).send({
    //     message: 'Rows deleted and modified data as JSON',
    //     data: modifiedSheetData,
    // });
    res.status(200).send({
        data: transformedResult,
    });
        
  },
};

function getDate(sheet) {    
    const dateCell = sheet[xlsx.utils.encode_cell({ r: 1, c: 0 })]; 
    const extractedText = dateCell ? dateCell.v : null;


    if (extractedText !== null) {
        const words = extractedText.split(' ');
        return words[words.length - 1];
    } else {
        console.log('the cell that is supposed to hold the date is empty');
    }
}

function modifyDate(date) {
    const monthAbbreviations = {
        "Jan": "01",
        "Feb": "02",
        "Mar": "03",
        "Apr": "04",
        "May": "05",
        "Jun": "06",
        "Jul": "07",
        "Agu": "08",
        "Sep": "09",
        "Oct": "10",
        "Nov": "11",
        "Dec": "12"
      };
    
      // Split the input date into parts
      const parts = date.split('-');
    
      // Extract the day, month abbreviation, and year
      const day = parts[0].replace('-', '');
      const monthAbbreviation = parts[1];
      const year = parts[2];
    
      // Convert the month abbreviation to a number
      const monthNumber = monthAbbreviations[monthAbbreviation];
    
      // Format the date in YYYY-MM-DD format
      const formattedDate = `${year}-${monthNumber}-${day.padStart(2, '0')}`;
    
      return formattedDate;
}

function delete4Rows(sheet) {
    // delete 4 rows
    function ec(r, c){
        return xlsx.utils.encode_cell({r:r,c:c});
    }

    var variable = xlsx.utils.decode_range(sheet["!ref"]);
    for (var R = variable.s.r; R < variable.s.r + 4; ++R) {
        for (var C = variable.s.c; C <= variable.e.c; ++C) {
            sheet[ec(R, C)] = "";
        }
    }
    variable.s.r += 4; // Update the start row to skip the first 4 rows
    variable.e.r -= 4; // Update the end row to account for the deleted rows
    sheet['!ref'] = xlsx.utils.encode_range(variable.s, variable.e);
    return sheet;    
}

function replaceEmptyWithNull(sheet) {
    const range = xlsx.utils.decode_range(sheet['!ref']);
  
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddress = { r: R, c: C };
        const cellValue = sheet[xlsx.utils.encode_cell(cellAddress)];
  
        if (cellValue === '') {
          sheet[xlsx.utils.encode_cell(cellAddress)] = null;
        }
      }
    }
  
    return sheet;
  }

