
const mainSearchInput = document.getElementById('mainSearchInput');
const downloadBtn = document.getElementById('downloadBtn');

const messageEl = document.getElementById('message');
const previewEl = document.getElementById('preview');

let currentPreview = null;

const handleSearch = async () => {
    const address = mainSearchInput.value.trim(); 
    if (!address) {
    messageEl.textContent = "Please enter an address or owner/management company.";
    return;
    }

    messageEl.textContent = "Searching...";
    previewEl.style.display = "none";

    try {
    const res = await fetch('/api/search-address', { 
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ address })
    });

    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'No compliance data found.');

    currentPreview = data;

    document.getElementById('previewAddress').textContent = data.Address || 'N/A';
    document.getElementById('previewOwner').textContent = data["Building Owner/Manager"] || 'N/A';
    document.getElementById('previewBorough').textContent = data.Borough || 'N/A';
    document.getElementById('previewFISP').textContent = data["FISP Compliance Status"] || 'N/A';

    previewEl.style.display = "block";
    messageEl.textContent = '';

    } catch (err) {
    messageEl.textContent = err.message;
    currentPreview = null;
    }
};

mainSearchInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
        handleSearch();
        e.preventDefault(); 
    }
});

downloadBtn.addEventListener('click', async () => {
    if (!currentPreview) {
    messageEl.textContent = "No data to generate report.";
    return;
    }

    messageEl.textContent = "Generating report...";

    try {
    const res = await fetch('/api/generate-report', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ address: currentPreview.Address })
    });

    if (!res.ok) {
        const errData = await res.json();
        throw new Error(errData.error || 'Failed to generate report.');
    }

    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Compliance_Report.docx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);

    messageEl.textContent = "Report downloaded successfully!";

    } catch (err) {
    messageEl.textContent = err.message;
    }
});
