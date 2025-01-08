const express = require("express");
const multer = require("multer");
const exceljs = require("exceljs");

const app = express();
const port = 8683;

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

const CELL_MAPPINGS = {
  checklistMasterName: {
    name: "Checklist Master Name",
    cell: "B1",
    maxLength: 255,
    required: true,
  },
  purpose: { name: "Purpose", cell: "B2", maxLength: 500, required: true },
  scopeOfUse: {
    name: "Scope of Use",
    cell: "B3",
    maxLength: 500,
    required: true,
  },
  usageFrequency: { name: "Usage Frequency", cell: "B4", required: false },
};

const CHECK_LIST_COLUMNS = {
  category: { name: "Category", col: "B", maxLength: 255 },
  item: { name: "Item", col: "C", maxLength: 255 },
  guideline: { name: "Guideline", col: "D", maxLength: 255 },
  required: { name: "Required", col: "E", maxLength: 255 },
};

function validateCell(value, config, cell, messages) {
  if (config.required && !value) {
    messages.push(`Cell ${cell}: ${config.name} cannot be empty!`);
  }
  if (value && config.maxLength && value.length > config.maxLength) {
    messages.push(`Cell ${cell}: ${config.name} is too long!`);
  }
}

function readBasicData(worksheet) {
  const data = {};
  const messages = [];

  for (const [key, config] of Object.entries(CELL_MAPPINGS)) {
    data[key] = worksheet.getCell(config.cell).value ?? "";
    validateCell(data[key], config, config.cell, messages);
  }

  return { data, messages };
}

function readCheckList(worksheet) {
  const checkList = [];
  const messages = [];
  let rowNumber = 7;

  while (true) {
    const row = worksheet.getRow(rowNumber);
    const rowData = {
      category: String(row.getCell(2).text ?? "").replace(/\n/g, " "),
      item: String(row.getCell(3).text ?? "").replace(/\n/g, " "),
      guideline: String(row.getCell(4).text ?? "").replace(/\n/g, " "),
      required: String(row.getCell(5).text ?? "").replace(/\n/g, " "),
    };

    if (!rowData.category && !rowData.item && !rowData.guideline) break;

    Object.entries(CHECK_LIST_COLUMNS).forEach(([field, config]) => {
      if (rowData[field].length > config.maxLength) {
        messages.push(
          `Cell ${config.col}${rowNumber}: ${config.name} is too long!`
        );
      }
    });

    checkList.push(rowData);
    rowNumber++;
  }

  return { checkList, messages };
}

function groupCheckListData(checkList) {
  const grouped = checkList.reduce((acc, entry) => {
    const key = entry.category.trim();
    if (!acc[key]) acc[key] = [];
    acc[key].push({
      item: entry.item,
      guideline: entry.guideline,
      required: entry.required,
      order: acc[key].length + 1,
    });
    return acc;
  }, {});

  return Object.keys(grouped).map((key, index) => ({
    category: key,
    items: grouped[key],
    order: index + 1,
  }));
}

app.post("/upload", upload.single("excelFile"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        status: 400,
        message: ["No file uploaded."],
        data: {},
      });
    }

    const workbook = new exceljs.Workbook();
    await workbook.xlsx.load(req.file.buffer);
    const worksheet = workbook.worksheets[0];
    // const sheetNames = workbook.worksheets.map((sheet) => sheet.name);
    // console.log(sheetNames);
    // const worksheet = workbook.getWorksheet(sheetNames[0]);
    
    const { data, messages: basicMessages } = readBasicData(worksheet);
    const { checkList, messages: checkListMessages } = readCheckList(worksheet);
    const allMessages = [...basicMessages, ...checkListMessages];

    if (allMessages.length === 0) {
      data.checkList = groupCheckListData(checkList).map((item) =>
        JSON.stringify(item)
      );

      return res.json({ status: 200, message: ["success"], data });
    }

    return res.status(400).json({
      status: 400,
      message: allMessages,
      data: {},
    });
  } catch (error) {
    console.error(error);
    return res.status(500).json({
      status: 500,
      message: ["An error occurred while processing the file."],
      data: {},
    });
  }
});

app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});



