const express=require('express')
const multer= require("multer")
const reader= require('xlsx')
const ExcelJS=require('exceljs');

const app=express()
const port=process.env.PORT || 3000


app.use(express.json())

app.use('/publicfiles',express.static(__dirname + '/publicfiles'))


app.get('/readexcelfile',(req,res)=>{
  let filename=req.query.filename;
  let data=[]
  try {
    const file=reader.readFile('publicfiles/' + filename + ".xlsx")
    const sheetNames=file.SheetNames

    for(let i=0;i<sheetNames.length;i++)
    {
      const arr= reader.utils.sheet_to_json(file.Sheets[sheetNames[i]])
      arr.forEach((res)=>{
        data.push(res)
      })
    }
    res.send(data);

  console.log(data)
  let i = 1;
  let co1_dis=0;
  let co1_fc=0;
  let co1_pass=0;

  let co2_dis=0;
  let co2_fc=0;
  let co2_pass=0;

  let co3_dis=0;
  let co3_fc=0;
  let co3_pass=0;

  let result={}
  for(i=0;i<256;i++)
  {
    
    if(data[i].CO1>=7)
    {
      co1_dis=co1_dis+1;
    }
    if(data[i].CO1>=6)
    {
      co1_fc=co1_fc+1;
    }
    if(data[i].CO1>=4)
    {
      co1_pass=co1_pass+1;
    }

    if(data[i].CO2>=7)
    {
      co2_dis=co2_dis+1;
    }
    if(data[i].CO2>=6)
    {
      co2_fc=co2_fc+1;
    }
    if(data[i].CO2>=4)
    {
      co2_pass=co2_pass+1;
    }

    if(data[i].CO3>=7)
    {
      co3_dis=co3_dis+1;
    }
    if(data[i].CO3>=6)
    {
      co3_fc=co3_fc+1;
    }
    if(data[i].CO3>=4)
    {
      co3_pass=co3_pass+1;
    }

    let i = 1;
    let co1_dis_perc=co1_dis*100/255;
    let co1_fc_perc=co1_fc*100/255;
    let co1_pass_perc=co1_pass*100/255;

    let co2_dis_perc=co2_dis*100/255;
    let co2_fc_perc=co2_fc*100/255;
    let co2_pass_perc=co2_pass*100/255;

    let co3_dis_perc=co3_dis*100/255;
    let co3_fc_perc=co3_fc*100/255;
    let co3_pass_perc=co3_pass*100/255;
    
    


  }



    console.log(co1_dis);
    console.log(co1_fc);
    console.log(co1_pass);  

    console.log(co2_dis);
    console.log(co2_fc);
    console.log(co2_pass);

    console.log(co3_dis);
    console.log(co3_fc);
    console.log(co3_pass);



  } catch (e) {
    res.send(e)
  }
})

var storage = multer.diskStorage({
    destination: 'publicfiles',
    filename: function (req, file, callback) {
        callback(null, file.originalname);
    }
});


const upload=multer({
  dest:'publicfiles',
  storage:storage,
  limits :{
    fileSize:1000000
  },
  fileFilter(req,file,cb){

    if(!file.originalname.match('xlsx')){
      return cb(new Error('Please upload an image'))
    }
    
    cb(undefined,true)
  }
})

const errorMiddleware =(req,res,next) =>{
    throw new Error('From my middleware')
} 

app.post('/upload',upload.single('upload'), async(req,res)=>{

  res.send()
},(error,req,res,next)=>{
   res.status(400).send({error:error.message})
})


app.post('/sheet',async(req,res)=>{
  try{
   // Requiring module
// const reader = require('xlsx')

// Reading our test file
let workbook=new ExcelJS.Workbook()
await workbook.xlsx.readFile('./test.xlsx')
let worksheet=workbook.getWorksheet("Sheet1")

let distinction = worksheet.getRow(3);
let firstclass= worksheet.getRow(4);
let secondclass = worksheet.getRow(5);

distinction.getCell(2).value = 2;
distinction.getCell(3).value = 2;
distinction.getCell(5).value = 2;
distinction.getCell(6).value = 2;

firstclass.getCell(2).value = 2;
firstclass.getCell(3).value = 2;
firstclass.getCell(5).value = 2;
firstclass.getCell(6).value = 2;


secondclass.getCell(2).value = 2;
secondclass.getCell(3).value = 2;
secondclass.getCell(5).value = 2;
secondclass.getCell(6).value = 2;


distinction.commit()
firstclass.commit()
secondclass.commit()

 await workbook.xlsx.writeFile('./test.xlsx');

    res.send('done');
  }
  catch(e){
    res.status(400).send(e);
  }
});


app.listen(port,()=>{
  console.log("Server is up on port "+ port)
})

