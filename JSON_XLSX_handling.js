let xlsx = require("xlsx");
let fs = require("fs");
// --------------------------  JSON  ------------------------
// METHOD-01
    // let buffer = fs.readFileSync("./example.json");
    // console.log(buffer);  // prints binary data.
    // let data = JSON.parse(buffer);
    // console.log(data)  // prints data as is stored in the json file.

// METHOD-02 (easier)
// var data = new Array;  //doesnt work because the array gets converted to obj
data = require("./example.json");  // ** must use the "[]" to convert from obj to arr IF JSON file is in obj format, i.e., if the outermost brackets of the JSON file is "{}" and not "[]".
console.log(data);
let pushItem = {
    "name": "Thor",
    "last Name": "-",  // space(" ") in key is only allowed when value is str.   ** this comment will get ignored and not be added to the JSON file.
    "isAvenger": true,
    "friends": ["Bruce", "Steve", "Natasha"],
    "age": 45,
    "address": {
        "dimension": 3,
        "place": "Asguard"
    }
};
data.push(pushItem);   // pushes data only in the current instance of file i.e. in RAM and not the actual file.
let stringData = JSON.stringify(data);  // converts data to str
fs.writeFileSync("example.json", stringData);  // writes the data in the file.


// --------------------------  XLSX  ------------------------
// FLOW OF XLSX FILES:
// File -> WorkBook -> WorkSheet -> Columns -> Rows
function excelWriter (filePath, jsonData, sheetName) {
    let newWB = xlsx.utils.book_new();
    let newWS = xlsx.utils.json_to_sheet(jsonData);
    xlsx.utils.book_append_sheet(newWB, newWS, sheetName);
    xlsx.writeFile(newWB, filePath);  //creates new if file is not present.
}
excelWriter("xlsx1.xlsx", data, "Sheet1");  //**** data must be PARSED JSON DATA not JSON FILE ****
function excelReader (filePath, sheetName) {
    if (!fs.existsSync(filePath)) {return [];}
    let wb = xlsx.readFile(filePath);
    let excelData = wb.Sheets[sheetName];
    let data = xlsx.utils.sheet_to_json(excelData);
    return data;
}
let excelData = excelReader("xlsx1.xlsx", "Sheet1");
console.log(excelData);