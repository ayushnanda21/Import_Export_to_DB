const express = require('express');
const fs = require('fs');
const multer = require("multer");
const bodyParser = require("body-parser")
const mongoose  =require("mongoose");
const path = require("path");
const userModel = require("./Models/userModel")
const excelToJson = require("convert-excel-to-json");


//multer storage
const storage = multer.diskStorage({
    destination:(req,file,cb)=>{  
        cb(null, './public/uploads');  
    },  
    filename:(req,file,cb)=>{  
        cb(null, file.originalname);  
    }  
});

const upload = multer({storage : storage});

//db
mongoose.connect('mongodb://localhost:27017/exceldemo',{useNewUrlParser:true})  
.then(()=>console.log('connected to db'))  
.catch((err)=>console.log(err)) 

//middlewares
const app =express();
app.use(bodyParser.urlencoded({extended: true}));
app.use(express.json());

//static folder
app.use(express.static('./public'));




//upload excel file
app.post('/uploadfile', upload.single("uploadfile"), (req, res) =>{
    importExcelData2MongoDB(__dirname  + '/public/uploads/'  + req.file.filename);
    res.status(200).json({
        'msg': 'File imported to database successfully',
        'file': req.file
    });

});



// Import Excel File to MongoDB database
function importExcelData2MongoDB(filePath, req){
    // -> Read Excel File to Json Data
    const excelData = excelToJson({
    sourceFile:  fs.readFileSync(filePath),
    sheets:[{
    // Excel Sheet Name
    name: 'demo',
    // Header Row -> be skipped and will not be present at our result object.
    header:{
    rows: 1
    },
    // Mapping columns to keys
    columnToKey: {
    A: 'name',
    B: 'email',
    C: 'age'
    }

    }]
});
    // -> Log Excel Data to Console
    console.log(filePath);
    console.log(excelData);



// //Insert jsonobject to mongodb
userModel.insertMany(excelData.demo, (err, res)=>{
    if(err){
        console.log(err);
    }

    console.log("Inserted Successfully");
});


}

//server
app.listen(5000, (req,res)=>{
    console.log("Server is running on port 5000");
})
