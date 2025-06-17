// accessing excel sheet using js promise one method ....

// const Exceljs = require('exceljs');
// const workbook = new Exceljs.Workbook();

// workbook.xlsx.readFile("F:/trend-notes/excelpractice.xlsx").then(() =>{
//     const worksheet = workbook.getWorksheet("Sheet1")
//     worksheet.eachRow((row,rownumber) =>{
//         row.eachCell((cell,colnumber) =>{
//             console.log(cell.value)
//         })
//     })
// })


//.....................................................
// accessing excel sheet using js async and await another method...

const Exceljs = require("exceljs")

async function readwrite(){
    let output = {row: -1, column: -1}
    const workbook = new Exceljs.Workbook();  
    await workbook.xlsx.readFile("F:/trend-notes/excelpractice.xlsx");
    const worksheet = workbook.getWorksheet("Sheet1")
    worksheet.eachRow((row,rownumber) =>{
        row.eachCell((cell,colnumber) =>{
           if(cell.value === "lakshmi"){
            output.row = rownumber
            output.column = colnumber
        console.log(rownumber,colnumber)
         }
        })
    })
}
let cell = worksheet.getCell(output.row,output.column)
cell.value = "lakshmi"
await workbook.xlsx.writeFile("F:/trend-notes/excelpractice.xlsx")
readwrite();

// accessing excel sheet using js async and await another method with col and row number (or) by the name ...

const Exceljs = require("exceljs")

async function readwrite(){
    const workbook = new Exceljs.Workbook();
    await workbook.xlsx.readFile("F:/trend-notes/excelpractice.xlsx");
    const worksheet = workbook.getWorksheet("Sheet1")
    worksheet.eachRow((row,rownumber) =>{
        row.eachCell((cell,colnumber) =>{
            //if(cell.value === "lakshmi"){
          //  console.log(cell.value)
            //}
        })
    })
    let cells = worksheet.getCell(2,6)
    console.log(cells.value)
}
readwrite();