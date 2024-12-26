const express = require('express');
const multer = require('multer');
const exceljs = require('exceljs');

const app = express();
const port = 8386;

// Cấu hình multer để lưu file tải lên vào bộ nhớ
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.post('/upload', upload.single('excelFile'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Không có file nào được tải lên.' });
    }

    const workbook = new exceljs.Workbook();
    // Đọc file từ bộ nhớ
    await workbook.xlsx.load(req.file.buffer);

    // Lấy worksheet đầu tiên
    const worksheet = workbook.getWorksheet(1);

    // Lấy giá trị của ô D1
    const cellValue = worksheet.getCell('D1').value;

    res.json({ data: cellValue });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Đã có lỗi xảy ra khi xử lý file.' });
  }
});

app.listen(port, () => {
  console.log(`Server đang lắng nghe tại cổng ${port}`);
});