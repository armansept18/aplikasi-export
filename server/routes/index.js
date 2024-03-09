const express = require("express");
const router = express.Router();
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");
const cors = require("cors");

router.get("/", function (req, res, next) {
  res.render("index", { title: "Express" });
});
router.get("/excel/aoa", cors(), (req, res) => {
  const header = ["ID", "Nama", "Umur"];
  const data = [
    { id: 1, name: "Denny", age: 24 },
    { id: 2, name: "Aditya", age: 25 },
    { id: 3, name: "Pradipta", age: 26 },
    { id: 4, name: "Ardhie", age: 27 },
    { id: 5, name: "Putra", age: 28 },
    { id: 6, name: "Prananta", age: 29 },
  ];
  const fileName = "AOA_XLS";
  let wb = XLSX.utils.book_new();
  wb.Props = {
    Title: fileName,
    Author: "Test",
    CreatedDate: new Date(),
  };
  wb.SheetNames.push("Sheet 1");
  let ws = XLSX.utils.json_to_sheet(data);
  wb.Sheets["Sheet 1"] = ws;

  const downloadFolder = path.resolve(__dirname, "../downloads");
  if (!fs.existsSync(downloadFolder)) {
    fs.mkdirSync(downloadFolder);
  }
  try {
    XLSX.writeFile(wb, `${downloadFolder}${path.sep}${fileName}.xls`);
    res.download(`${downloadFolder}${path.sep}${fileName}.xls`);
  } catch (e) {
    console.log(e.message);
    throw e;
  }
});
module.exports = router;
