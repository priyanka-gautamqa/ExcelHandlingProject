const ExcelJs = require('exceljs');


//readFile is a async step two ways to handle  - use await with async or use promise or use then()

/**
 * const workbook = new ExcelJs.Workbook(); //created the Object of class Workbook using ExcelJs

 * workbook.xlsx.readFile('ExcelDownloadTest.xlsx').then(function(){

    const worksheet = workbook.getWorksheet('Sheet1');

//print all values of the excel
worksheet.eachRow((row,rowNumber)=>{
    row.eachCell((cell,colNumber)=>{
        console.log(cell.value);
    })
})
})
*/

//OTHER WAY
async function excelTest(){

    const workbook = new ExcelJs.Workbook(); 
    await workbook.xlsx.readFile('ExcelDownloadTest.xlsx');
    const worksheet = workbook.getWorksheet('Sheet1');
    let output = {row:-1,column:-1};

    //print all values of the excel
    worksheet.eachRow((row,rowNumber)=>{
        row.eachCell((cell,colNumber)=>{
            //console.log(cell.value); -  to print all values
            if(cell.value==='Apple'){
                console.log(rowNumber,colNumber)
                output.row=rowNumber;
                output.column=colNumber;
            }
        })
    })

    //replace some cell value
    const cell = worksheet.getCell(output.row,output.column);
    cell.value = 'PEACH';

    //save the above change 
    workbook.xlsx.writeFile('ExcelDownloadTest.xlsx');

}
excelTest();