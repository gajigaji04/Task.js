const express = require("express");
const multer = require("multer");
const XlsxPopulate = require("xlsx-populate");
const { Sequelize, DataTypes } = require("sequelize");

const app = express();
const port = 3000;

// 뷰 엔진으로 EJS 설정
app.set("view engine", "ejs");
app.set("views", __dirname + "/views");

// Sequelize 설정
const sequelize = new Sequelize({
  dialect: "sqlite",
  storage: "database.sqlite",
});

// 모델 정의
const Data = sequelize.define("Data", {
  customerId: {
    type: DataTypes.INTEGER,
    allowNull: false,
  },
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

app.post("/upload_excel", upload.single("file"), async (req, res) => {
  try {
    const buffer = req.file.buffer;

    // xlsx-populate를 사용하여 Excel 파일 읽기
    const workbook = await XlsxPopulate.fromDataAsync(buffer);
    const sheet = workbook.sheet(0);

    // 데이터베이스에 엑셀 데이터 등록
    const usedRange = sheet.usedRange();
    const rowCount = usedRange.endCell().rowNumber();

    for (let row = 2; row <= sheet.rowCount; row++) {
      const customerId = sheet.cell(`A${row}`).value();
      const name = sheet.cell(`B${row}`).value();
      const customerGrade = sheet.cell(`C${row}`).value();

      console.log(
        `Row ${row} - Customer ID: ${customerId}, Name: ${name}, Grade: ${customerGrade}`
      );

      if (name !== null && customerGrade !== null) {
        await Data.create({ name, value: customerGrade });
      } else {
        console.log(`Skipping row ${row} due to null values.`);
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
