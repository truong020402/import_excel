const express = require("express");
const multer = require("multer");
const exceljs = require("exceljs");

const app = express();
const port = 8386;

// Cấu hình multer để lưu file tải lên vào bộ nhớ
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.post("/upload", upload.single("excelFile"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "Không có file nào được tải lên." });
    }

    const workbook = new exceljs.Workbook();
    // Đọc file từ bộ nhớ
    await workbook.xlsx.load(req.file.buffer);

    // Lấy worksheet đầu tiên
    const worksheet = workbook.getWorksheet("Checklist Master");

    const data = {};

    // Read specific cells based on their coordinates
    data.checklistMasterName = worksheet.getCell("D1").value ?? "";
    console.log(data.checklistMasterName);
    data.purpose = worksheet.getCell("N1").value ?? "";
    console.log(data.purpose);
    data.scopeOfUse = worksheet.getCell("D2").value ?? "";
    console.log(data.scopeOfUse);
    data.usagePeriodFrom = worksheet.getCell("H2").value ?? "";
    console.log(data.usagePeriodFrom);
    data.usagePeriodTo = worksheet.getCell("N2").value ?? "";
    console.log(data.usagePeriodTo);
    data.submissionAddress = worksheet.getCell("D3").value ?? "";
    console.log(data.submissionAddress);
    data.usageFrequency = worksheet.getCell("H3").value ?? "";
    console.log(data.usageFrequency);
    data.usageFrequencyNotes = worksheet.getCell("N3").value ?? "";
    console.log(data.usageFrequencyNotes);
    data.searchTags = worksheet.getCell("D4").value ?? "";
    console.log(data.searchTags);

    // Bắt đầu đọc dữ liệu từ hàng thứ 2 (bỏ qua hàng tiêu đề)
    let rowNumber = 9;
    let currentRow = worksheet.getRow(rowNumber);
    const Check_list = [];

    const Category = 2; // Cột B
    const Item = 3; // Cột C
    const Guideline = 4; // Cột D


    while (currentRow.getCell(Category).value || currentRow.getCell(Item).value || currentRow.getCell(Guideline).value) {
      const checkList = {
        category: currentRow.getCell(Category).value,
        item: currentRow.getCell(Item).value,
        guideline: currentRow.getCell(Guideline).value,
      };
      Check_list.push(checkList);
      rowNumber++;
      currentRow = worksheet.getRow(rowNumber);

    }
    data.checkList = Check_list;
    res.json({ data });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Đã có lỗi xảy ra khi xử lý file." });
  }
});

app.listen(port, () => {
  console.log(`Server đang lắng nghe tại cổng ${port}`);
});
