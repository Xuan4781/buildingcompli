import express from 'express';
import 'dotenv/config';
import cors from 'cors';
import path, { dirname, join } from 'path';
import fs from 'fs';
import xlsx from 'xlsx';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

let buildingData = [];

// Load Excel data
function loadExcelData() {
    try {
        const workbook = xlsx.readFile(join(__dirname, "final_merged.xlsx"));
        const sheetName = workbook.SheetNames[0];
        buildingData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { cellDates: true });
        console.log(`✅ Excel data loaded: ${buildingData.length} records`);
        return true;
    } catch (err) {
        console.error("Failed to load Excel:", err.message);
        return false;
    }
}

// Compliance logic
function determineCompliance(excelData){
    const fispStatus = excelData['FISP Compliance Status'];
    const fispLastFilingStatus = excelData['FISP Last Filing Status'];
    if(fispStatus === "UNSAFE" || fispStatus === "No Report Filed") return 'Non-Compliant';
    if(fispStatus === "SWARMP" || fispLastFilingStatus?.includes('SWARMP')) return 'In Compliance';
    const fispDue = new Date(excelData['FISP Filing Due']);
    if(fispDue < new Date()) return 'Non-Compliant';
    return 'In Compliance';
}

// Map Excel
function mapExcel(excelData){
    return {
        'Address': excelData['Address'],
        'Building_OwnerManager': excelData['Building Owner/Manager'],
        'Borough': excelData['Borough'],
        'FISP Compliance Status': determineCompliance(excelData),
        'Contact Email': excelData['Contact Email'] || 'Chelsea.Coppinger@socotec.us',
        'Contact Phone': excelData['Contact Phone'] || '+1 646 549 6045',
        // Add other fields...
    };
}

// API: refresh data
app.get('/api/refresh-data', (req, res) => {
    if (loadExcelData()) res.send('Excel refreshed');
    else res.status(500).send('Failed to refresh Excel');
});

// API: search
app.post('/api/search-address', (req, res) => {
    const { address } = req.body;
    if (!address) return res.status(400).json({ error: 'Address required' });
    const searchAddress = address.trim().toLowerCase();
    const found = buildingData.find(r => r.Address?.toLowerCase() === searchAddress);
    if (found) res.json(found);
    else res.status(404).json({ error: 'Address not found' });
});

// API: generate report
app.post('/api/generate-report', (req, res) => {
    const { address } = req.body;
    if (!address) return res.status(400).json({ error: 'Address required' });

    const reportData = buildingData.find(r => r.Address?.toLowerCase() === address.trim().toLowerCase());
    if (!reportData) return res.status(404).json({ error: 'Data not found' });

    const mappedData = mapExcel(reportData);

    // Process values
    const processedData = {};
    for (const [key, value] of Object.entries(mappedData)) {
        const str = String(value).trim().toLowerCase();
        processedData[key] = (value == null || str === '' || str === 'undefined' || str === 'null' || str === 'nan' || Number.isNaN(value)) ? 'N/A' : value;
    }

    try {
        const templatePath = join(__dirname, "newtemp.docx");
        const content = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, delimiters: { start: '«', end: '»' } });
        doc.render(processedData);
        const buf = doc.getZip().generate({ type: 'nodebuffer' });
        res.setHeader('Content-Disposition', 'attachment; filename=Compliance_Report.docx');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.send(buf);
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Failed to generate report' });
    }
});

// SPA fallback
app.get('*', (req, res) => {
    res.sendFile(join(__dirname, 'public', 'index.html'));
});

// Start server
const PORT = process.env.PORT || 5001;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
    loadExcelData();
});
