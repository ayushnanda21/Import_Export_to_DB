

const mongoose = require("mongoose");

const excelSchema = new mongoose.Schema(
{
     name:{
         type: String
     },
     email:{
         type: String
     },
     age:{
         type: Number
     }

});

module.exports = mongoose.model("userModel", excelSchema);