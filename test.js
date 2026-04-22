const XLSX = require('xlsx');

const exportData = [
    { "Email": "test@test.com", "BSc Grade": 10, "MSc Grade": 10, "Research Exp.": 10, "Prof. Exp.": 10, "English Skills": 10, "CV & Cover Letter": 10, "Total Score": "" }
];
const ws = XLSX.utils.json_to_sheet(exportData);

const weights = { bsc: 4, msc: 32, research: 16, prof: 8, english: 8, cv: 32 };

for (let i = 0; i < exportData.length; i++) {
    const rowNum = i + 2;
    const formula = `B${rowNum}*(${weights.bsc}/100)+C${rowNum}*(${weights.msc}/100)+D${rowNum}*(${weights.research}/100)+E${rowNum}*(${weights.prof}/100)+F${rowNum}*(${weights.english}/100)+G${rowNum}*(${weights.cv}/100)`;
    const cellRef = XLSX.utils.encode_cell({c: 7, r: i + 1});
    ws[cellRef] = { t: 'n', f: formula };
}

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "Grades");
XLSX.writeFile(wb, "test_grades.xlsx");
console.log("Written test_grades.xlsx");
