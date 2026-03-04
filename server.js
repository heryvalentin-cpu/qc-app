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
app.use(express.static("public"));

// Upload folder
if (!fs.existsSync("uploads")){
    fs.mkdirSync("uploads");
}

// MySQL Connection
const db = mysql.createConnection({
    host: "localhost",
    user: "root",
    password: "changeme",
    database: "qc_app"
});

db.connect(err => {
    if (err) throw err;
    console.log("MySQL Connected...");
});

// Multer config
const storage = multer.diskStorage({
    destination: "./uploads/",
    filename: (req, file, cb) => {
        cb(null, Date.now() + "-" + file.originalname);
    }
});

const upload = multer({ storage });

// Insert QC data
app.post("/add", upload.single("file"), (req, res) => {

    const { reference, description, qc_name, status } = req.body;

    const file_path = req.file ? req.file.path : null;

    db.query(
        "INSERT INTO checks (reference, description, file_path, qc_name, status) VALUES (?,?,?,?,?)",
        [reference, description, file_path, qc_name, status],
        err => {
            if (err) throw err;
            res.send("Data inserted");
        }
    );
});

// Import CSV / Excel structure variable
app.post("/import", upload.single("file"), (req, res) => {

    if (!req.file) return res.send("No file uploaded");

    let data = [];

    const filePath = req.file.path;

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
        const sheet = workbook.SheetNames[0];
        const sheetData = workbook.Sheets[sheet];

        data = xlsx.utils.sheet_to_json(sheetData);
    }

    // Insert dynamic data
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

// Export CSV
app.get("/export/csv", (req, res) => {

    db.query("SELECT * FROM checks", (err, results) => {

        const parser = new Parser();
        const csv = parser.parse(results);

        res.header("Content-Type", "text/csv");
        res.attachment("checks.csv");
        res.send(csv);
    });
});

// Export Excel
app.get("/export/excel", (req, res) => {

    db.query("SELECT * FROM checks", async (err, results) => {

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("QC Data");

        if (results.length > 0){
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
            "attachment; filename=qc_data.xlsx"
        );

        await workbook.xlsx.write(res);
        res.end();
    });
});

app.listen(3000, () => console.log("Server running on port 3000"));