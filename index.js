import express from "express";
import mysql from "mysql2/promise";
import cors from "cors";
import ExcelJS from "exceljs";

// Create a connection pool
const pool = mysql.createPool({
  host: "nextsoftwarethailand.com",
  user: "nextsoft_dev_01",
  password: "nextsoft1234",
  database: "nextsoft_dev_01",
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0,
});

const app = express();
app.use(express.json());
app.use(cors());

app.get("/", (req, res) => {
  res.json({ message: "555555555555" });
});

app.get("/cars/", async (req, res) => {
  const search = req.query.search;
  try {
    let data = "";

    if (search === "") {
      const sql = "SELECT * FROM LPR";
      const result = await pool.query(sql);
      data = result[0];
    } else {
      const sql = "SELECT * FROM LPR WHERE lp LIKE ?";
      const result = await pool.query(sql, [`%${search}%`]);
      data = result[0];
    }

    res.status(200).json({
      message: "success",
      data,
    });
  } catch (error) {
    console.log(error);
  }
});

app.get("/excel", async (req, res) => {
  try {
    const sql = "SELECT * FROM LPR";
    const result = await pool.query(sql);

    // Create a new Excel workbook
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Sheet 1");

    // Add headers to the worksheet
    const headers = Object.keys(result[0]);
    worksheet.addRow(headers);

    // Add data to the worksheet
    result[0].forEach((row) => {
      const values = Object.values(row);
      worksheet.addRow(values);
    });

    // Save the workbook to a buffer
    workbook.xlsx
      .writeBuffer()
      .then((buffer) => {
        // Set response headers for Excel file download
        res.setHeader(
          "Content-Disposition",
          "attachment; filename=exported_data.xlsx"
        );
        res.setHeader(
          "Content-Type",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );
        res.send(buffer);
        // connection.end(); // Close the database connection
      })
      .catch((writeError) => {
        console.error("Error writing Excel buffer:", writeError);
        res.status(500).send("Internal Server Error");
        // connection.end(); // Close the database connection
      });
  } catch (error) {
    console.log(error);
  }
});

app.listen(3000, () => {
  console.log("server is 300");
});
