const express = require("express");
const multer = require("multer");
const XlsxPopulate = require("xlsx-populate");
const { Sequelize, DataTypes } = require("sequelize");

const app = express();
const port = 3000;

// Sequelize 설정
const sequelize = new Sequelize({
  dialect: "sqlite",
  storage: "database.sqlite",
});

// 모델 정의
const Data = sequelize.define("Data", {
  name: {
    type: DataTypes.STRING,
    allowNull: false,
  },
  value: {
    type: DataTypes.FLOAT,
    allowNull: false,
  },
});

// 미들웨어 설정
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.set("view engine", "ejs");
app.set("views", __dirname + "/views");

app.get("/", (req, res) => {
  res.render("index");
});

app.post("/upload_excel", upload.single("file"), async (req, res) => {
  try {
    const buffer = req.file.buffer;

    // Use xlsx-populate to read Excel file
    const workbook = await XlsxPopulate.fromDataAsync(buffer);
    const sheet = workbook.sheet(0);

    // 데이터베이스에 엑셀 데이터 등록
    for (let row = 2; row <= sheet.rowCount(); row++) {
      const name = sheet.cell(`A${row}`).value();
      const value = sheet.cell(`B${row}`).value();

      await Data.create({ name, value });
    }

    res.json({ message: "Excel data uploaded successfully" });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Internal Server Error" });
  }
});

// 서버 시작
sequelize.sync().then(() => {
  app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
  });
});
