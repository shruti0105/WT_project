const ExcelJS=require('exceljs')

const writeExcelFile= async function(UA,co_final,subject)
{
  try{

    console.log(UA)
    console.log(co_final)
   // Requiring module
// const reader = require('xlsx')

// Reading our test file
let workbook=new ExcelJS.Workbook()
await workbook.xlsx.readFile('./Output.xlsx')
let worksheet=workbook.getWorksheet("Sheet1")

let co= worksheet.getRow(5);
let firstclass= worksheet.getRow(6);
let secondclass = worksheet.getRow(7);

co.getCell(2).value =  co_final.CO1;
co.getCell('B5').numFmt = '0.0000000000';
co.getCell(3).value = co_final.CO2;
co.getCell('C5').numFmt = '0.000000000';
co.getCell(4).value = co_final.CO3;
co.getCell(5).value = co_final.CO4;
co.getCell(6).value = co_final.CO5;
co.getCell(7).value =co_final.CO6;
let avg=(co_final.CO1+co_final.CO2+co_final.CO3+co_final.CO4+co_final.CO5+co_final.CO6)/6
co.getCell(8).value =avg

co.getCell(9).value = UA.resultd;
co.getCell(10).value = UA.targetd;

let dis=0;
if((UA.resultd/UA.targetd)*3>3)
  dis=3
else
  dis=(UA.resultd/UA.targetd)*3

firstclass.getCell(9).value = UA.resultf;
firstclass.getCell(10).value = UA.targetf;

let fir=0;
if((UA.resultf/UA.targetf)*2>2)
  fir=2
else
  fir=(UA.resultf/UA.targetf)*2


secondclass.getCell(9).value = UA.results;
secondclass.getCell(10).value = UA.targets;

let sec=0;
if((UA.results/UA.targets)>1)
  sec=1
else
  sec=(UA.results/UA.targets)

co.commit()
firstclass.commit()
secondclass.commit()


let d= worksheet.getRow(10)
let f= worksheet.getRow(11)
let p= worksheet.getRow(12)
let a= worksheet.getRow(13)

d.getCell(9).value = dis
f.getCell(9).value = fir
p.getCell(9).value = sec

let avg2=(dis+fir+sec)/6
a.getCell(9).value = avg2

d.commit()
f.commit()
p.commit()
a.commit()

let attain=worksheet.getRow(15)
attain.getCell(6).value=avg*0.3
attain.getCell(9).value=avg2*0.7

attain.commit()

let final=worksheet.getRow(17)
final.getCell(7).value=avg*0.3+avg2*0.7
final.commit() 

let name="Final " +subject+ " attainment AY 2022-23";

 await workbook.xlsx.writeFile('./'+name+ '.xlsx');

    // res.send('done');
  }
  catch(err){
  console.log(err)
  }
}

module.exports=writeExcelFile