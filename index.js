// Requiring module
const excel = require("exceljs");
const express = require("express");
const bodyParser = require("body-parser");
const fs = require("fs");
const util = require("util");
const unlinkFile = util.promisify(fs.unlink);
const app = express();
const multer = require("multer");
const port = 3000;

app.use(bodyParser.urlencoded({ extended: false }));

// parse application/json
app.use(bodyParser.json());

const upload = multer({ dest: "./public/data/uploads/" });

const readExcelFile = async (filename) => {
  var workbook = new excel.Workbook();
  const resSet = [];

  await workbook.xlsx.readFile(filename).then(function () {
    //get sheet name from the excel file
    workbook.eachSheet(({ name }) => {
      let sheetArray = [];
      var worksheet = workbook.getWorksheet(name);
      worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
        //   console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
        sheetArray.push(row.values);
      });

      resSet.push([...sheetArray]);
    });
  });

  console.log(resSet);
  await unlinkFile(filename);
  return resSet;
};

app.post("/", upload.single("file"), async (req, res) => {
  const file = req.file;
  const result = await readExcelFile(file.path);
  //   console.log(file);
  res.send(result);
});

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`);
});
