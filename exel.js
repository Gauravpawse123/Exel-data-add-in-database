const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const mysql = require('mysql2');
const path = require('path');
const app = express();

app.use(express.json());
const port = 8000;

// Set up Multer storage
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/');
    },
    filename: (req, file, cb) => {
        cb(null, Date.now() + path.extname(file.originalname));
    }
});

const upload = multer({ storage: storage });

// MySQL database connection
const db = mysql.createConnection({
    host: 'localhost',
    user: 'root',
    password: '',
    database: 'exel'
});

db.connect((err) => {
    if (err) throw err;
    console.log('Connected to MySQL database.');
});


app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html');
});


app.post('/upload', upload.single('excel'), (req, res) => {
    const filePath = req.file.path;


    const workbook = xlsx.readFile(filePath);
    const sheet_name_list = workbook.SheetNames;
    const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

    data.forEach(row => {
        const sql = 'INSERT INTO data SET ?';
        db.query(sql, row, (err, result) => {
            if (err) throw err;
            console.log('Data inserted:', result.insertId);
        });
    });

    res.send('File uploaded and data inserted into the database.');
});

app.listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
});
