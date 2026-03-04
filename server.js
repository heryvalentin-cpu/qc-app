require("dotenv").config();
const express = require("express");
const mysql = require("mysql2");
const multer = require("multer");
const cors = require("cors");
const fs = require("fs");
const xlsx = require("xlsx");
const { Parser } = require("json2csv");
const ExcelJS = require("exceljs");

const app = express();

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static("public"));

// Create uploads folder if not exists
if (!fs.existsSync("uploads")) {
    fs.mkdirSync("uploads");
}

// MySQL Cloud Connection
const db = mysql.createConnection({
    host: process.env.DB_HOST,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB_NAME,
    port: 17331
});

db.connect(err => {
    if (err) {
        console.error("MySQL Connection Error:", err);
        return;
    }
    console.log("MySQL Connected...");
});

// Multer Storage Config
const storage = multer.diskStorage({
    destination: "./uploads/",
    filename: (req, file, cb) => {
        cb(null, Date.now() + "-" + file.originalname);
    }
});

const upload = multer({ storage });

/*
====================================
 QC DATA INSERT
====================================
*/
app.post("/add", upload.single("file"), (req, res) => {

    const { reference, description, qc_name, status } = req.body;

    const file_path = req.file ? req.file.path : null;

    const sql = `
    INSERT INTO checks (reference, description, file_path, qc_name, status)
    VALUES (?,?,?,?,?)
    `;

    db.query(sql,
        [reference, description, file_path, qc_name, status],
        err => {
            if (err) {
                console.error(err);
                return res.status(500).send("Database insert error");
            }

            res.send("Data inserted");
        });
});

/*
====================================
 IMPORT CSV / EXCEL VARIABLE STRUCTURE
====================================
*/
app.post("/import", upload.single("file"), (req, res) => {

    if (!req.file) return res.send("No file uploaded");

    const filePath = req.file.path;
    let data = [];

    if (req.file.originalname.endsWith(".csv")) {

        const content = fs.readFileSync(filePath, "utf8");
        const rows = content.split("\n");
        const headers = rows[0].split(",");

        for (let i = 1; i < rows.length; i++) {

            if (!rows[i]) continue;

            const values = rows[i].split(",");

            let rowData = {};

            headers.forEach((header, index) => {
                rowData[header.trim()] = values[index];
            });

            data.push(rowData);
        }

    } else {

        const workbook = xlsx.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        data = xlsx.utils.sheet_to_json(sheet);
    }

    data.forEach(row => {

        const columns = Object.keys(row);

        if (columns.length === 0) return;

        const sql = `
        INSERT INTO checks (${columns.join(",")})
        VALUES (${columns.map(() => "?").join(",")})
        `;

        db.query(sql, Object.values(row));
    });

    res.send("Import successful");
});

/*
====================================
 EXPORT CSV
====================================
*/
app.get("/export/csv", (req, res) => {

    db.query("SELECT * FROM checks", (err, results) => {

        const parser = new Parser();
        const csv = parser.parse(results);

        res.header("Content-Type", "text/csv");
        res.attachment("qc_export.csv");
        res.send(csv);
    });
});

/*
====================================
 EXPORT EXCEL
====================================
*/
app.get("/export/excel", (req, res) => {

    db.query("SELECT * FROM checks", async (err, results) => {

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("QC Data");

        if (results.length > 0) {
            worksheet.columns = Object.keys(results[0]).map(key => ({
                header: key,
                key: key
            }));

            worksheet.addRows(results);
        }

        res.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );

        res.setHeader(
            "Content-Disposition",
            "attachment; filename=qc_export.xlsx"
        );

        await workbook.xlsx.write(res);
        res.end();
    });
});

/*
====================================
 SERVER START
====================================
*/
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
    console.log("Server running on port", PORT);
});