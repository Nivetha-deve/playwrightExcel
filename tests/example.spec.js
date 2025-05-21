// task 
// for all data in excel to check and pass sauclab...
import { test,expect } from "@playwright/test";

const Exceljs = require('exceljs');

async function sauclab(){ 
const workbook = new Exceljs.Workbook();
await workbook.xlsx.readFile("F:/trend-notes/sauclab.xlsx");
const worksheet = workbook.getWorksheet("Sheet1");

let user = []

worksheet.eachRow((row,rownumber) =>{
    if(rownumber > 1){
         const username = row.getCell(1).value;
         const password = row.getCell(2).value;
         user.push({username,password});
    }
})
return user;
}

test("login sauglab", async({page}) =>{
    let userData = await sauclab();
    for(const user of userData){
    await page.goto("https://www.saucedemo.com/v1/")
    await page.locator('[id="user-name"]').fill(user.username);
    await page.locator('[id="password"]').fill(user.password);
    await page.getByText("LOGIN").click();
    
    }
})

// to check one username and password check practice page...
