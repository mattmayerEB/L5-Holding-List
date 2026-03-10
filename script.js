/**
 * Retention Tracking Form - connects to Excel via CSV export
 * Column A: Agent | B: Loan ID | C: Quote ID | D: Customer Name | E: Cancel Reason | F: Cancel Fee | G: Submitted Date
 */

const STORAGE_KEY = 'retentionEntries';
const FLOW_URL = 'https://script.google.com/macros/s/AKfycbx8F9uD2Wu7V6TXGebNtKjYRgQqiqunf7HgyEfeWT2YNdUoLCgoEalGKK6UBNijk4RJ/exec';
// Load entries from localStorage
function getEntries() {
    const data = localStorage.getItem(STORAGE_KEY);
    return data ? JSON.parse(data) : [];
}

// Save entries to localStorage
function saveEntries(entries) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(entries));
}

// Add entry and refresh UI
function addEntry(entry) {
    const entries = getEntries();
    entries.push(entry);
    saveEntries(entries);
    renderEntries();
}

// Render entries list
function renderEntries() {
    const entries = getEntries();
    const listEl = document.getElementById('entriesList');
    const countEl = document.getElementById('entryCount');

    countEl.textContent = entries.length;

    if (entries.length === 0) {
        listEl.innerHTML = '<p class="empty-state">No entries yet. Submit the form to add entries.</p>';
        return;
    }

    listEl.innerHTML = entries.map((e) => `
        <div class="entry-item">
            <strong>${e.loanId}</strong> — ${e.user || '—'} | Quote: ${e.quoteId || '—'} | ${e.customerName} | ${e.cancelReason} | Fee: ${e.cancelFee || '—'} | Date: ${e.submittedAt || '—'}
        </div>
    `).join('');
}

// Export entries to CSV (opens in Excel)
function exportToCSV() {
    const entries = getEntries();

    if (entries.length === 0) {
        alert('No entries to export. Add entries first.');
        return;
    }

    const headers = ['Agent', 'Loan ID', 'Quote ID', 'Customer Name', 'Cancel Reason', 'Cancel Fee', 'Submitted Date'];
    const rows = entries.map(e => [
        e.user || '',
        e.loanId,
        e.quoteId || '',
        e.customerName,
        e.cancelReason,
        e.cancelFee || '',
        e.submittedAt || ''
    ]);

    const csvContent = [
        headers.join(','),
        ...rows.map(row => row.map(cell => {
            const s = String(cell);
            return s.includes(',') || s.includes('"') ? `"${s.replace(/"/g, '""')}"` : s;
        }).join(','))
    ].join('\n');

    const blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `retention_entries_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
}

// Form submit handler
document.getElementById('retentionForm').addEventListener('submit', async function (e) {
    e.preventDefault();

    const today = new Date();
    const submittedAt = today.toISOString().slice(0, 10); // YYYY-MM-DD

    const entry = {
        user: document.getElementById('user').value,
        loanId: document.getElementById('loanId').value.trim(),
        quoteId: document.getElementById('quoteId').value.trim(),
        customerName: document.getElementById('customerName').value.trim(),
        cancelReason: document.getElementById('cancelReason').value,
        cancelFee: document.getElementById('cancelFee').value.trim(),
        submittedAt
    };

    try {
        const response = await fetch(FLOW_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'text/plain'
            },
            body: JSON.stringify(entry)
        });

        const text = await response.text();
        if (!response.ok) {
            throw new Error('Server returned ' + response.status + ': ' + text.slice(0, 100));
        }

        addEntry(entry);
        this.reset();
        alert('Submitted successfully.');
    } catch (err) {
        console.error(err);
        const msg = err.message || err.toString();
        const hint = msg.includes('Failed to fetch') || msg.includes('NetworkError')
            ? '\n\nTip: Open the form via a local server instead of double-clicking the file. In this folder run: npx serve'
            : '';
        alert('Submit failed: ' + msg + hint);
    }
});

document.getElementById('exportBtn').addEventListener('click', exportToCSV);

// Initial render
renderEntries();
