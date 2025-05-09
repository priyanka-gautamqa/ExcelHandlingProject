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

async function writeExcelTest(searchText,newValue,fileName){

    const workbook = new ExcelJs.Workbook(); 
    await workbook.xlsx.readFile(fileName);
    const worksheet = workbook.getWorksheet('Sheet1');
  

    const output = await readExcel(worksheet,searchText);

    //replace some cell value
    const cell = worksheet.getCell(output.row,output.column);
    cell.value = newValue;

    //save the above change 
    await workbook.xlsx.writeFile(fileName);

}


async function readExcel(worksheet,searchText){
    //print all values of the excel
    let output = {row:1,column:1};
    worksheet.eachRow((row,rowNumber)=>{
        row.eachCell((cell,colNumber)=>{
           // console.log(cell.value); //-  to print all values
            if(cell.value===searchText){
                output.row=rowNumber;
                output.column=colNumber;
                console.log("rowNumber",rowNumber);
            }
        })
    })
    return output;
}

//send searchText from here
writeExcelTest('Apple','PEACHES','ExcelDownloadTest.xlsx')