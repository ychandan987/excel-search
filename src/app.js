const express = require('express');
const hbs = require('hbs');
const cors = require('cors')
// const ApiResponse = require('./utils/ApiResponse.js');

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const port = process.env.PORT || 3001

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

app.use(express.json());
app.use(express.urlencoded());
app.use(cors());



app.get('',(req,res) =>{
    res.render('index.hbs',{
        title : 'Excel Report',
        name : 'Gem Fashion Studio'
    })
})



app.get('/excel_report', async (req,res) => {
    const test = req.query.address;
    console.log(test);
    
    if(!test){
        return res.send({error:'You must provide an values'})
    }

    // Define the folder path containing Excel files
        const folderPath = "./My Files/";

    // Define the value you are searching for
        const searchValue = test;

    // Function to search for a value in an Excel sheet
        function searchInSheet(sheet, searchValue) {
            try {
                const range = xlsx.utils.decode_range(sheet['!ref']);
                // console.log(range)
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
            } catch (error) {
                return null;
            }
        }

        // Function to search through all Excel files in the folder
        async function searchInFolder(folderPath, searchValue) {
            return new Promise((resolve, reject)=>{
                let result = [];
                fs.readdir(folderPath, (err, files) => {
                    if (err) {
                        console.error("Could not list the directory.", err);
                        return;
                    }

                    for (const file of files) {
                        const filePath = path.join(folderPath, file);
                        console.log(filePath);
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
                                    result.push({
                                        searchValue,
                                        filePath,
                                        sheetName,
                                        cellRef
                                    });
                                    // console.log(`Found "${searchValue}" in file: ${filePath}, sheet: ${sheetName}, cellRef: ${cellRef}`);
                                    
                                    //    return res.json({ excelReport: test_data });
                                }
                            
                            });
                        }
                    }
                    resolve(result);
                });
            });
        
        }
            
        // Run the search
        const response = await searchInFolder(folderPath, searchValue);
        res.send(response);

    
})

app.listen(port,() => {
    console.log(`Server is up on the port ${port}`)
})