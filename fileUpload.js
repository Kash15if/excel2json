// Requiring module
const excel = require("exceljs");
const express = require("express");
const bodyParser = require("body-parser");
const fs = require("fs");
const util = require("util");
const unlinkFile = util.promisify(fs.unlink);
const { uuid } = require("uuidv4");
const app = express();
const multer = require("multer");
const port = 3000;

app.use(bodyParser.urlencoded({ extended: false }));

// parse application/json
app.use(bodyParser.json());

var storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "./public/data/uploads/");
  },
  filename: function (req, file, cb) {
    // cb(null, uuid());
    cb(null, "abcd");
  },
});

const upload = multer({ storage: storage });

const readExcelFile = async (filename) => {
  var workbook = new excel.Workbook();
  const resSet = [];

  await workbook.xlsx.readFile(filename).then(function () {
    var worksheet = workbook.getWorksheet("sheet1");
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      //   console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
      resSet.push(JSON.stringify(row.values));
    });
  });

  await unlinkFile(filename);
  return resSet;
};

app.post("/", upload.single("file"), async (req, res) => {
  const file = req.file;
  const result = await readExcelFile(file.path);
  console.log(file);
  res.send(result);
});

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`);
});
