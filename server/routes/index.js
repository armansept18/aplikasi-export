const express = require("express");
const router = express.Router();
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");
const cors = require("cors");

var mysql = require('mysql');

var con = mysql.createConnection({
    host: "127.0.0.1",
    user: "root",
    password: "password",
    database: "direct-debit" // Add database name here

});

con.connect(function (err) {
    if (err) throw err;
    console.log("Connected!");
});

function excelGenerator(fileName, author, sheetName, data) {
    const wb = XLSX.utils.book_new();
    wb.Props = {
        Title: fileName,
        Author: author,
        CreatedDate: new Date(),
    };
    wb.SheetNames.push(sheetName);
    wb.Sheets[sheetName] = XLSX.utils.json_to_sheet(data);

    const downloadFolder = path.resolve(__dirname, "../downloads");
    if (!fs.existsSync(downloadFolder)) {
        fs.mkdirSync(downloadFolder);
    }

    XLSX.writeFile(wb, `${downloadFolder}${path.sep}${fileName}.xls`);

    return `${downloadFolder}${path.sep}${fileName}.xls`
}

router.get("/", function (req, res, next) {
    res.render("index", {title: "Express"});
});
router.get("/excel/payments", cors(), (req, res) => {
    // const users = db.users.find({})
    con.query('SELECT * FROM payments', function (err, result) {
        if (err) throw err;
        console.log(result);

        const data = result.map(payment => {
            return {
                'ID': payment.id,
                'Nomor Referensi': payment.reference_no,
                'Token Pembayaran': payment.charge_token,
            }
        })

        const newExcel = excelGenerator('payments-jan', 'dbs', 'January', data)

        res.download(newExcel);
    });
});

router.get("/excel/cards", cors(), (req, res) => {
    con.query('SELECT * FROM cards', function (err, result) {
        if (err) throw err;
        console.log(result);

        const data = result.map(card => {
            return {
                'ID': card.id,
                'Card Data': card.card_data,
            }
        })

        const newExcel = excelGenerator('cards-jan', 'dbs', 'January', data)

        res.download(newExcel);
    });
});

module.exports = router;
