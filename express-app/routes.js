const express = require("express"),
  router = express.Router();

const fs = require("fs");
const xlsxToJSON = require("xlsx-to-json");
const JSZip = require('jszip');
const Docxtemplater = require("docxtemplater");
const mkdirp = require("mkdirp"); 
const multer = require("multer");
  
const multerConfig = {
    
  storage: multer.diskStorage({
   //Setup where the user's file will go
   destination: function(req, file, next){
     next(null, './public/files');
     },   
      
      //Then give the file a unique name
      filename: function(req, file, next){
          next(null, file.originalname);
        }
      }),   
    };

function readJSONFile(filename, callback) {
  fs.readFile(filename, function (err, data) {
    if(err) {
      callback(err);
      return;
    }
    try {
      callback(null, JSON.parse(data));
    } catch(exception) {
      callback(exception);
    }
  });
}

//GET home page.
router.get("/", function(req, res) {
  res.render('index');
})

router.post("/", multer(multerConfig).single('excelFile'),  function(req, res) {
  try {
     xlsxToJSON({
         input: req.file.path,
         output: "input.json",
         // lowerCaseHeaders: true
       },
       function(err, result) {
         if (err) {
           res.end();
         } else {
           res.redirect("/generateWordFile")
         }
       }
     );
      } catch(err) {
        res.redirect("/error")
      }
});

router.get("/generateWordFile", function(req, res) {
readJSONFile('input.json', function(err, json) {
  if(err) res.render('error');
  res.render('generateFile', {input: json});
})
})

router.post("/generateWordFile", multer(multerConfig).single('templateFile'), function(req, res) {

  readJSONFile('input.json', function(err, json) {
      if(err) {
          throw err;
      }
      let saveToPath = require('path').join(require('os').homedir(), 'Desktop');
      let folder = saveToPath + `/${req.body.folderName}`;
      mkdirp(folder, function(err) { 
          if(err) throw err;
      });

      try {
        let counter = 0;
      for (let obj of json) {
         let content = fs.readFileSync(req.file.path, "binary");

          let zip = new JSZip(content);
         
           let doc = new Docxtemplater();
           doc.loadZip(zip);
          
           //set the templateVariables
           doc.setData(obj);
          
           try {
             // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
             doc.render();
           } catch (error) {
             var e = {
               message: error.message,
               name: error.name,
               stack: error.stack,
               properties: error.properties
             };
             console.log(JSON.stringify({ error: e }));
             // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
             throw error;
           }
          
           var buf = doc.getZip().generate({ type: "nodebuffer" });
          
           // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
           
           fs.writeFileSync(`${saveToPath}/${req.body.folderName}/${Object.values(obj)[1]}-${counter}.docx`, buf);
           counter++;
      }
      res.redirect("/generatedSuccessfully");
    }catch(err) {
      res.redirect("/error");
    }
  });
});

router.get("/generatedSuccessfully", function(req, res) {
res.render('generatedSuccessfully.hbs');
})

router.get("/error", function(req, res) {
res.render('error.hbs');
})

router.get("*", function(req, res) {
res.status(400).render('error.hbs');
})

module.exports = router;
