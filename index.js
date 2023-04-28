const express=require('express')
const multer= require("multer")
const reader= require('xlsx')
const ExcelJS=require('exceljs');

const app=express()
const port=process.env.PORT || 3000

let co_final={};
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

    let count=data.length-1;

  // console.log(data);

  // ut1_boolean=[1,1,1,0,0,0];
    // CO_char=["CO1","CO2","CO3","CO4","CO5","CO6"];
    // let count=256
    // let co_dis=0;
    // let co_fc=0;
    // let co_pass=0;

    
    // for(let i=0;i<6;i++)
    // { 
    
    // let ut1_boolean=[1,1,1,0,0,0];
    
    console.log(data)
    for (var key in data[0]) {

      
    let co1_dis=0;
    let co1_fc=0;
    let co1_pass=0;
    let total1=0
      if(key=="Roll No. / Max marks" || key=="Total")
      // console.log(data[0][key])
      {
        continue;
      }
      console.log(key);

        for(let i=0;i<count;i++)
    {
      
        if(data[i][key]>=7)
        {
          co1_dis=co1_dis+1;
        }
        if(data[i][key]>=6)
        {
          co1_fc=co1_fc+1;
        }
        if(data[i][key]>=4)
        {
          co1_pass=co1_pass+1;
        }
      
if(data[i][key]!=0 && data[i][key]!="")
    {
      total1=total1+1;
    }
  }
    let high_target=50;
    let mid_target=63 ;
    let low_target=82;
    let co1_dis_perc;
    let co1_fc_perc;
    let co1_pass_perc;
    
    co1_dis_perc=co1_dis*100/total1;
    co1_fc_perc=co1_fc*100/total1;
    co1_pass_perc=co1_pass*100/total1;
    
    let co1_att_highl=(co1_dis_perc/high_target)*3;
    if(co1_att_highl>3)
    {
      co1_att_highl=3;
    }
    else{
      co1_att_highl=co1_dis_perc/high_target;
    }
    let co1_att_midl=(co1_fc_perc/mid_target)*2;
    if(co1_att_midl>2)
    {
      co1_att_midl=2;
    }
    else{
      co1_att_highl=co1_dis_perc/high_target*2;
    }
    let co1_att_lowl=(co1_pass_perc/low_target);
    if(co1_att_lowl>1)
    {
      co1_att_lowl=1;
    }
    else{
      co1_att_highl=co1_dis_perc/high_target;
    }
    let co1_attainment=(co1_att_highl+co1_att_midl+co1_att_lowl)/6;
    const hasKey = key in co_final;
    if(hasKey)
    {
      co_final[key]=(co_final[key]+co1_attainment)/2;
    }
    else{
      co_final[key]=co1_attainment;
    }
  //   console.log(co1_dis);
  //   console.log(co1_fc);
  //   console.log(co1_pass);
  //   console.log(co1_dis_perc);
  //   console.log(co1_fc_perc);
  //   console.log(co1_pass_perc);
  //   console.log(co1_att_highl);
  //  console.log(co1_att_midl);
  //  console.log(co1_att_lowl);
  //  console.log(co1_attainment);

      
}

for (var key in co_final)
{
  console.log(key+"-->"+co_final[key]);
  
} 
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
