const express=require('express')
const multer= require("multer")
const hbs=require('hbs')
const  path = require('path')
const compute=require("./compute.js")
const writeExcelFile=require("./excel.js")
const upload=require("./uploadFile.js")

const app=express()
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


// console.log(path.join(__dirname,'../views/pdf.hbs'))

app.get('',(req,res)=>{
  res.render('home')
})


// const errorMiddleware =(req,res,next) =>{
//     throw new Error('From my middleware')
// } 

app.post('/submit',upload.array('upload',2), async(req,res)=>{
  
  //  res.render('generate')
  console.log(req.files[0].filename)
  //  console.log(req.files[1].filename)
  console.log(req.body)
  let co_final={
  };
  let UA={
    resultd:req.body.resultd,
    targetd:req.body.targetd,
    resultf:req.body.resultf,
    targetf:req.body.targetf,
    results:req.body.results,
    targets:req.body.targets
  }
  let ut1={
     CO1:req.body.CO1ut1,
     CO2:req.body.CO2ut1,
     CO3:req.body.CO3ut1,
     CO4:req.body.CO4ut1,
     CO5:req.body.CO5ut1,
     CO6:req.body.CO6ut1,
  }
  let ut2={
     CO1:req.body.CO1ut2,
     CO2:req.body.CO2ut2,
     CO3:req.body.CO3ut2,
     CO4:req.body.CO4ut2,
     CO5:req.body.CO5ut2,
     CO6:req.body.CO6ut2,
  }
  compute(req.body.uth1,req.body.utm1,req.body.utl1,co_final,req.files[0].filename,req.body.no1,ut1)
  compute(req.body.uth2,req.body.utm2,req.body.utl2,co_final,req.files[1].filename,req.body.no2,ut2)
  writeExcelFile(UA,co_final,req.body.subject)

   const options = {
        root: path.join(__dirname ,"../")
    };
 
    // const fileName = 'Final.xlsx';
    // res.sendFile(fileName, options, function (err) {
    //     if (err) {
    //         next(err);
    //     } else {
    //         console.log('Sent:', fileName);
    //     }
    // });
  // compute(req.body.uth2,req.body.utm2,req.body.utl2,co_final,req.files[1].filename)
},(error,req,res,next)=>{
   res.status(400).send({error:error.message})
})


// app.get('/downloadExcel', function (req, res,next) {
  
// });

// app.get('/downloadExcel', (req, res, next) => {
//   const excelFilePath = path.join(__dirname, '../Final.xlsx');
//   console.log(excelFilePath)
//   res.sendFile(excelFilePath, (err) => {
//     if (err) console.log(err);
//   });
// });
app.listen(port,()=>{
  console.log("Server is up on port "+ port)
})