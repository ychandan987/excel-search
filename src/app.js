
const express = require('express')
const hbs = require('hbs')
// const ApiResponse = require('./utils/ApiResponse.js');

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { log } = require('console');

const port = process.env.PORT || 3000

// console.log(__dirname)
// console.log(path.join(__dirname,'../public'))
const pathPublic = path.join(__dirname,'../public')
const viewPath = path.join(__dirname, '../templates')
const partialPath = path.join(__dirname, '../templates/partials')

const app = express()
app.set('view engine','hbs')
app.set('views', viewPath);
hbs.registerPartials(partialPath)
app.use(express.static(pathPublic))

app.use(express.json({limiit :"16kb"}));
app.use(express.urlencoded({extended : true, limit : "16kb"}));



app.get('',(req,res) =>{
    res.render('index.hbs',{
        title : 'Excel Report',
        name : 'Gem Fashion Studio'
    })
})



app.get('/excel_report',(req,res) => {
    const test = req.query.address;
    console.log(test);
    
    if(!test){
        return res.send({error:'You must provide an values'})
    }

    // Define the folder path containing Excel files
        const folderPath = "A:/chandan/Project/excel-search/My Files/";

    // Define the value you are searching for
        const searchValue = test;

    // Function to search for a value in an Excel sheet
        function searchInSheet(sheet, searchValue) {
            const range = xlsx.utils.decode_range(sheet['!ref']);
            for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
                for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
                    const cellAddress = { c: colNum, r: rowNum };
                    const cellRef = xlsx.utils.encode_cell(cellAddress);
                    const cell = sheet[cellRef];
                    if (cell && cell.v == searchValue) {
                        return cellRef;
                    }
                }
            }
            return null;
        }

    // Function to search through all Excel files in the folder
            function searchInFolder(folderPath, searchValue) {
                fs.readdir(folderPath, (err, files) => {
                    if (err) {
                        console.error("Could not list the directory.", err);
                        return;
                    }

                    files.forEach((file) => {
                        const filePath = path.join(folderPath, file);
                        if (filePath.endsWith(".xlsx") || filePath.endsWith(".xls") || filePath.endsWith("/.xlsx")) {
                            const workbook = xlsx.readFile(filePath);
                            workbook.SheetNames.forEach(sheetName => {
                                const sheet = workbook.Sheets[sheetName];
                                const cellRef = searchInSheet(sheet, searchValue);
                                
                                if (cellRef) {

                                    var test_data = {
                                        searchValue : searchValue,
                                        file : file,
                                        sheetName : sheetName,
                                        cellRef : cellRef
                                    };
                                    
                                                                        
                                    // console.log(test_data);
                                    // return test_data;
                                    console.log(`Found "${searchValue}" in file: ${filePath}, sheet: ${sheetName}, cellRef: ${cellRef}`);
                                    
                                    //    return res.json({ excelReport: test_data });
                                }
                            
                            });
                        }
                    });
                });
            }
            
        // Run the search
        searchInFolder(folderPath, searchValue);

    
})

app.listen(port,() => {
    console.log(`Server is up on the port ${port}`)
}) 

