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

// Load Excel data into memory
function loadExcelData() {
    try {
        const workbook = xlsx.readFile(join(__dirname, "final_merged.xlsx"));
        const sheetName = workbook.SheetNames[0];
        buildingData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { cellDates: true });
        console.log(`✅ Excel data loaded. ${buildingData.length} records in memory.`);
        return true;
    } catch (error) {
        console.error("Failed to load 'final_merged.xlsx'.", error.message);
        return false;
    }
}

// Determine FISP compliance
function determineCompliance(excelData){
    const fispStatus = excelData['FISP Compliance Status'];
    const fispLastFilingStatus = excelData['FISP Last Filing Status'];

    if(fispStatus === "UNSAFE" || fispStatus === "No Report Filed") return 'Non-Compliant';
    if(fispStatus === "SWARMP" || fispLastFilingStatus?.includes('SWARMP')) return 'In Compliance';

    const currentDate = new Date();
    const fispDue = new Date(excelData['FISP Filing Due']);
    if(fispDue < currentDate) return 'Non-Compliant';

    return 'In Compliance';
}

// Map Excel row for template
function mapExcel(excelData){
    const complianceStatus = determineCompliance(excelData);
    return {
        'Address': excelData['Address'],
        'Building_OwnerManager': excelData['Building Owner/Manager'],
        'Use Type': excelData['Use Type'],
        'Block': excelData['Block'],
        'BIN': excelData['BIN'],
        'Borough': excelData['Borough'],
        'Year Built': excelData['Year Built'],
        'M Floors': excelData['M Floors'],
        'Approx_Sq_Ft': excelData['Approx Sq Ft'],
        'Landmark': excelData['Landmark'],
        'Parking Garage (Yes/No)': excelData['Parking Garage'],
        //'FISP Compliance Status': excelData['FISP Compliance Status'],
        'Sub': excelData['Sub'],
        'FISP Filing Due': excelData['FISP Filing Due'],
        'FISP Last Filing Status': excelData['FISP Last Filing Status'],
        'FISP Cycle Filing Window': excelData['FISP Cycle Filing Window'],
        'LL126 Compliance Status': excelData['LL126 Compliance Status'],
        'LL126 Cycle': excelData['LL126 Cycle'],
        'LL126 Previous Filing Status': excelData['LL126 Previous Filing Status'],
        'LL126 SREM Recommended Date': excelData['LL126 SREM Recommended Date'],
        'LL126 Filing Window': excelData['LL126 Filing Window'],
        'LL126 Filing Due': excelData['LL126 Filing Due'],
        'LL126 Next Steps': excelData['LL126 Next Steps'],
        'LL126 Parapet Compliance Status': excelData['LL126 Parapet Compliance Status'],
        'LL84 Compliance Status': excelData['LL84 Compliance Status'],
        'LL84 Filing Due': excelData['LL84 Filing Due'],
        'LL84 Next Steps': excelData['LL84 Next Steps'],
        'LL87 Compliance Status': excelData['LL87 Compliance Status'],
        'LL87 Filing Due': excelData['LL87 Filing Due'],
        'LL87 Compliance Year': excelData['LL87 Compliance Year'],
        'LL87 Next Steps': excelData['LL87 Next Steps'],
        'LL88 Compliance Status': excelData['LL88 Compliance Status'],
        'LL88 Filing Due': excelData['LL88 Filing Due'],
        'LL88 Notes': excelData['LL88 Notes'],
        'LL97 Compliance Status': excelData['LL97 Compliance Status'],
        'LL97 Filing Due': excelData['LL97 Filing Due'],
        'LL97 Next Steps': excelData['LL97 Next Steps'],
        'Contact Email': excelData['Contact Email'] || 'Chelsea.Coppinger@socotec.us',
        'Contact Phone': excelData['Contact Phone'] || '+1 646 549 6045',
    };
}

// Refresh Excel data
app.get('/api/refresh-data', (req, res) => {
    if (loadExcelData()) res.status(200).send('Excel data refreshed.');
    else res.status(500).send('Failed to refresh Excel data.');
});

// Search by address
app.post('/api/search-address', (req, res) => {
    const { address } = req.body;
    if (!address) return res.status(400).json({ error: 'Address is required.' });

    const searchAddress = address.trim().toLowerCase();
    const found = buildingData.find(
        row => row.Address && row.Address.toString().trim().toLowerCase() === searchAddress
    );

    if (found) res.json(found);
    else res.status(404).json({ error: 'Address not found in Excel.' });
});

// Generate Word report
app.post('/api/generate-report', (req, res) => {
    const { address } = req.body;
    if (!address) return res.status(400).json({ error: 'Address is required.' });

    const reportData = buildingData.find(
        row => row.Address && row.Address.toString().trim().toLowerCase() === address.trim().toLowerCase()
    );
    if (!reportData) return res.status(404).json({ error: 'Data not found.' });

    const mappedData = mapExcel(reportData);

    // --- PROCESS DATA: replace undefined/null/empty/NAN values with 'N/A' ---
    const processedData = {};
    for (const [key, value] of Object.entries(mappedData)) {
        let processedValue = value;
        const stringValue = String(value).trim().toLowerCase();

        if (
            value === undefined ||
            value === null ||
            value === '' ||
            stringValue === '' ||
            stringValue === 'undefined' ||
            stringValue === 'null' ||
            stringValue === 'nan' ||
            Number.isNaN(value)
        ) {
            processedValue = 'N/A';
        }
        processedData[key] = processedValue;
    }

    try {
        const templatePath = join(__dirname, "newtemp.docx");
        const content = fs.readFileSync(templatePath, "binary");
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, delimiters: { start: '«', end: '»' } });

        // Render using processedData instead of mappedData
        doc.render(processedData);

        const buf = doc.getZip().generate({ type: "nodebuffer" });

        res.setHeader("Content-Disposition", "attachment; filename=Compliance_Report.docx");
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        res.send(buf);

    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Failed to generate report.' });
    }
});

// Generate Word report
app.post('/api/generate-report', (req, res) => {
    const { address } = req.body;
    if (!address) return res.status(400).json({ error: 'Address is required.' });

    const reportData = buildingData.find(
        row => row.Address && row.Address.toString().trim().toLowerCase() === address.trim().toLowerCase()
    );
    if (!reportData) return res.status(404).json({ error: 'Data not found.' });

    const mappedData = mapExcel(reportData);

    try {
        const templatePath = join(__dirname, "newtemp.docx");
        const content = fs.readFileSync(templatePath, "binary");
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, delimiters: { start: '«', end: '»' } });
        doc.render(mappedData);
        const buf = doc.getZip().generate({ type: "nodebuffer" });

        res.setHeader("Content-Disposition", "attachment; filename=Compliance_Report.docx");
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        res.send(buf);

    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Failed to generate report.' });
    }
});
app.use(express.static('public'));


// Start server
const PORT = process.env.PORT || 5001;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
    loadExcelData();
});
