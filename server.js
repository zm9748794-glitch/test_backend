const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 5000;
const EXCEL_FILE = path.join(__dirname, 'submissions.xlsx');
const ITEMS_FILE = path.join(__dirname, 'items.json');
const LOG_FILE = path.join(__dirname, 'log.txt'); // ملف اللوج الجديد

// Middleware
app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// Load items from JSON file
function loadItems() {
  try {
    const data = fs.readFileSync(ITEMS_FILE, 'utf8');
    return JSON.parse(data);
  } catch (error) {
    console.error('Error loading items:', error);
    return { categories: [] };
  }
}

// Save items to JSON file
function saveItems(items) {
  try {
    fs.writeFileSync(ITEMS_FILE, JSON.stringify(items, null, 2));
    console.log('✅ Items inventory updated successfully');
  } catch (error) {
    console.error('❌ Error saving items:', error);
  }
}

// Append logs to log.txt
function writeLog(entry) {
  try {
    fs.appendFileSync(LOG_FILE, entry + '\n', 'utf8');
    console.log('📝 Log entry saved');
  } catch (error) {
    console.error('❌ Error writing to log.txt:', error);
  }
}

// Initialize Excel file with ALL columns
async function initializeExcelFile() {
  if (!fs.existsSync(EXCEL_FILE)) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Registrations');

    worksheet.columns = [
      { header: 'Timestamp', key: 'timestamp', width: 22 },
      { header: 'Full Name', key: 'name', width: 25 },
      { header: 'Email', key: 'email', width: 32 },
      { header: 'Phone', key: 'phone', width: 18 },
      { header: 'Governorate', key: 'governorate', width: 20 },
      { header: 'Position in Team', key: 'position', width: 22 },
      { header: 'Committee', key: 'committee', width: 25 },
      { header: 'Category', key: 'category', width: 20 },
      { header: 'Item Booked', key: 'item', width: 25 },
      { header: 'Notes', key: 'notes', width: 40 }
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF6366F1' } };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow.height = 25;
    headerRow.eachCell((cell) => {
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
    });

    await workbook.xlsx.writeFile(EXCEL_FILE);
    console.log('✅ Excel file created');
  } else {
    console.log('✅ Excel file already exists');
  }
}

// GET items
app.get('/items', (req, res) => {
  try {
    const items = loadItems();
    res.json({ success: true, data: items });
  } catch (error) {
    console.error('❌ Error fetching items:', error);
    res.status(500).json({ success: false, message: 'Failed to fetch items' });
  }
});

// POST submit
app.post('/submit', async (req, res) => {
  try {
    const { name, email, phone, governorate, position, committee, notes, category, item } = req.body;

    if (!name || !email || !phone || !governorate || !position || !committee || !category || !item) {
      return res.status(400).json({ success: false, message: 'Missing required fields' });
    }

    const itemsData = loadItems();
    const categoryObj = itemsData.categories.find(cat => cat.name === category);
    if (!categoryObj) return res.status(404).json({ success: false, message: 'Category not found' });

    const itemObj = categoryObj.items.find(itm => itm.name === item);
    if (!itemObj) return res.status(404).json({ success: false, message: 'Item not found' });

    if (itemObj.count <= 0) {
      return res.status(400).json({ success: false, message: 'Item out of stock' });
    }

    const previousCount = itemObj.count;
    itemObj.count -= 1;
    saveItems(itemsData);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);
    const worksheet = workbook.getWorksheet('Registrations');

    const timestamp = new Date().toLocaleString('en-US', {
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true
    });

    const newRow = worksheet.addRow({
      timestamp, name, email, phone, governorate,
      position, committee, category, item, notes
    });
    newRow.commit();
    await workbook.xlsx.writeFile(EXCEL_FILE);

    // Write to log.txt
    const logEntry = `[${timestamp}] ${name} (${email}, ${phone}) - ${category}/${item} | Governorate: ${governorate}, Position: ${position}, Committee: ${committee}, Notes: ${notes || 'None'} | Stock: ${previousCount} → ${itemObj.count}`;
    writeLog(logEntry);

    res.json({ success: true, message: 'Booking successful', data: itemsData });
  } catch (error) {
    console.error('❌ ERROR:', error);
    res.status(500).json({ success: false, message: 'Server error: ' + error.message });
  }
});

// Health check
app.get('/health', (req, res) => {
  res.json({ status: 'Server running', port: PORT, time: new Date().toISOString() });
});

// Start server
app.listen(PORT, async () => {
  await initializeExcelFile();
  console.log(`🚀 Server running at http://localhost:${PORT}`);
});
