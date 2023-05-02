const express=require('express')
const multer= require("multer")
const ExcelJS=require('exceljs')
const hbs=require('hbs')
const  path = require('path')
const compute=require("./compute.js")

const app=express()
let co_final={};
const port=process.env.PORT || 3000


app.use(express.json())

app.use('/publicfiles',express.static(__dirname + '/publicfiles'))
app.use(express.static(path.join(__dirname,'../public')))
//define paths for expess config

const viewsPAth=path.join(__dirname,'../views') 
const bodyParser = require('body-parser')

//setup handlebars engine and view location.
app.set('view engine','hbs')
app.set('views',viewsPAth)


//setup static directory to serve

// app.use(express.static(path.join(__dirname,'../public'))) 
app.use(bodyParser.urlencoded({ extended: false }));


console.log(path.join(__dirname,'../views/pdf.hbs'))

app.get('',(req,res)=>{
  res.render('home')
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
      return cb(new Error('Please upload an xlsx'))
    }
    
    cb(undefined,true)
  }
})

const errorMiddleware =(req,res,next) =>{
    throw new Error('From my middleware')
} 

app.post('/submit',upload.array('upload',2), async(req,res)=>{
  
  const h=req.body.uth
  const m=req.body.utm
  const l=req.body.utl
  compute(h,m,l,co_final)
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

// app.post('/submit',async(req,res)=>{
  
//   const subject=req.body
//   console.log(subject)
// })
app.listen(port,()=>{
  console.log("Server is up on port "+ port)
})