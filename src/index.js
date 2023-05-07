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
    CO4:0.997881355932203,
    CO5:0.997881355932203,
    CO6:0.940775254924579,
  };
  let UA={
    resultd:req.body.resultd,
    targetd:req.body.targetd,
    resultf:req.body.resultf,
    targetf:req.body.targetf,
    results:req.body.results,
    targets:req.body.targets
  }
  compute(req.body.uth1,req.body.utm1,req.body.utl1,co_final,req.files[0].filename)
  writeExcelFile(UA,co_final,req.body.subject)

  const options = {
    root: path.join(__dirname ,"../")
};

const fileName = 'Final.xlsx';
res.sendFile(fileName, options, function (err) {
    if (err) {
        next(err);
    } else {
        console.log('Sent:', fileName);
    }
});
  // compute(req.body.uth2,req.body.utm2,req.body.utl2,co_final)
},(error,req,res,next)=>{
   return res.status(400).send({error:error.message})
})


// app.get('/downloadExcel', function (req, res,next) {
//     const options = {
//         root: path.join(__dirname ,"../")
//     };
 
//     const fileName = 'Final.xlsx';
//     res.sendFile(fileName, options, function (err) {
//         if (err) {
//             next(err);
//         } else {
//             console.log('Sent:', fileName);
//         }
//     });
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