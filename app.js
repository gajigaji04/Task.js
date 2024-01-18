const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { Sequelize, DataTypes } = require("sequelize");

const app = express();
const port = 3000;

// Set 'ejs' as the view engine
app.set("view engine", "ejs");

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

app.get("/", (req, res) => {
  res.render("index");
});

// 엑셀 파일 업로드 및 데이터베이스에 저장
app.post("/upload_excel", upload.single("file"), async (req, res) => {
  try {
    const buffer = req.file.buffer;
    const workbook = XLSX.read(buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    // 데이터베이스에 엑셀 데이터 등록
    for (const cellAddress in sheet) {
      if (cellAddress[0] === "A" && parseInt(cellAddress.slice(1)) > 1) {
        const rowNumber = parseInt(cellAddress.slice(1));
        const name = sheet[`A${rowNumber}`].v;
        const value = sheet[`B${rowNumber}`].v;

        await Data.create({ name, value });
      }
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
