const multer= require("multer")

var storage = multer.diskStorage({
    destination: 'publicfiles',
    filename: function (req, file, callback) {
        // console.log(file.originalname)
        callback(null, file.originalname);
    }
});


const upload=multer({
  dest:'publicfiles',
  storage:storage,
  filename:storage.filename,
  // filename: function(req, file, cb) {
  //   cb(null, file.originalname );
  // },
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


module.exports=upload