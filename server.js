const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const port = process.env.PORT || 3000; // Port para sa iyong server

// Middleware para tanggapin ang JSON body mula sa client
app.use(express.json());

// I-serve ang iyong static files (HTML, CSS, client-side JS) mula sa 'public' folder
app.use(express.static(path.join(__dirname, 'public')));

// Endpoint para sa Excel export
app.post('/export-excel', async (req, res) => {
    const { inventoryData, dateShift, timeCounted, shiftType } = req.body;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Countsheet');

    // --- Customize ang Design at Format ng Excel File Dito ---

    // Halimbawa: Setting up headers with specific styling
    worksheet.addRow(['PACKAGING MATERIALS DAILY COUNT SHEET - COHIN (BLDG. 3&6)']);
    worksheet.mergeCells('A1:D1');
    worksheet.getCell('A1').font = { name: 'Roboto Slab', size: 14, bold: true };
    worksheet.getCell('A1').alignment = { horizontal: 'center' };
    worksheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } }; // Light blue-gray

    worksheet.addRow([]); // Blank row for spacing

    worksheet.addRow(['Date / Shift :', dateShift]);
    worksheet.getCell('A3').font = { name: 'Roboto Slab', bold: true };
    worksheet.getCell('B3').font = { name: 'Courier New', italic: true, bold: true, underline: 'single' };

    worksheet.addRow(['Shift:', shiftType]);
    worksheet.getCell('A4').font = { name: 'Roboto Slab', bold: true };
    worksheet.getCell('B4').font = { name: 'Courier New', italic: true, bold: true, underline: 'single' };

    worksheet.addRow(['TIME COUNTED:', timeCounted]);
    worksheet.getCell('A5').font = { name: 'Roboto Slab', bold: true };
    worksheet.getCell('B5').font = { name: 'Courier New', italic: true, bold: true, underline: 'single' };

    worksheet.addRow([]); // Blank row

    // Table Headers
    const tableHeaders = ["ITEM CODE", "ACTUAL COUNT BREAKDOWN", "TOTAL", "REMARKS"];
    const headerRow = worksheet.addRow(tableHeaders);
    headerRow.eachCell((cell) => {
        cell.font = { name: 'Roboto Slab', bold: true, size: 10 };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
            top: { style: 'thin' }, bottom: { style: 'thin' },
            left: { style: 'thin' }, right: { style: 'thin' }
        };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } }; // Tailwind's slate-200
    });

    // Column Widths
    worksheet.columns = [
        { key: 'itemCode', width: 25 },
        { key: 'countBreakdown', width: 40 },
        { key: 'total', width: 15 },
        { key: 'remarks', width: 30 }
    ];

    // Data Rows
    inventoryData.forEach(item => {
        const total = calculateTotal(item.count);
        const excelFormula = generateExcelFormula(item.count);

        const row = worksheet.addRow([
            item.code,
            item.count.join(' | '),
            { formula: `=${excelFormula}`, result: total }, // Excel formula cell
            item.remarks.join(' | ')
        ]);

        // Apply borders to data cells
        row.eachCell((cell) => {
            cell.border = {
                top: { style: 'thin' }, bottom: { style: 'thin' },
                left: { style: 'thin' }, right: { style: 'thin' }
            };
            cell.alignment = { vertical: 'middle' };
        });

        // Example: Conditional Formatting based on remarks
        if (item.remarks.some(r => r.toLowerCase().includes("hold"))) {
            row.eachCell((cell) => {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfde68a' } }; // yellow-300
            });
        } else if (item.remarks.some(r => r.toLowerCase().includes("approved"))) {
            row.eachCell((cell) => {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBCFE8' } }; // pink-300
            });
        } else if (item.remarks.some(r => r.toLowerCase().includes("first out") || r.toLowerCase().includes("old"))) {
            row.eachCell((cell) => {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFa7f3d0' } }; // green-300
            });
        }
    });

    // --- Helper Functions (Copy from your client-side, or refine as needed) ---
    function calculateTotal(countArr) {
        if (!Array.isArray(countArr)) return 0;
        let total = 0;
        countArr.forEach(part => {
            const trimmedVal = String(part).trim(); // Ensure part is a string
            if (trimmedVal === "") return;

            if (trimmedVal.includes('×')) {
                const parts = trimmedVal.split('×');
                if (parts.length === 2) {
                    const num1 = parseFloat(parts[0].trim());
                    const num2 = parseFloat(parts[1].trim());
                    if (!isNaN(num1) && !isNaN(num2)) {
                        total += (num1 * num2);
                    }
                }
            } else if (trimmedVal.includes('+')) {
                const sumParts = trimmedVal.split('+').map(p => parseFloat(p.trim()));
                const sum = sumParts.reduce((acc, val) => acc + (isNaN(val) ? 0 : val), 0);
                total += sum;
            } else if (trimmedVal.includes('-')) {
                const subParts = trimmedVal.trim().split('-');
                if (subParts.length > 0) {
                    const initialValue = parseFloat(subParts[0].trim());
                    if (!isNaN(initialValue)) {
                        let subTotal = initialValue;
                        for (let i = 1; i < subParts.length; i++) {
                            const val = parseFloat(subParts[i].trim());
                            if (!isNaN(val)) {
                                subTotal -= val;
                            }
                        }
                        total += subTotal;
                    }
                }
            } else {
                const n = parseFloat(trimmedVal);
                if (!isNaN(n)) {
                    total += n;
                }
            }
        });
        return total;
    }

    function generateExcelFormula(countArr) {
        if (!Array.isArray(countArr) || countArr.length === 0) {
            return "0";
        }
        const parsedParts = countArr.map(part => {
            const trimmedVal = String(part).trim(); // Ensure part is a string
            if (trimmedVal === "") {
                return "0";
            }
            let formattedPart = trimmedVal.replace(/×/g, '*');

            if (formattedPart.includes('+') || formattedPart.includes('-')) {
                return `(${formattedPart})`; // Wrap sums/subtractions in parentheses
            }
            return formattedPart;
        });
        return parsedParts.join('+');
    }

    // --- Finalize at Ipadala ang File ---
    const dateParts = dateShift.split('/');
    const monthIndex = parseInt(dateParts[0], 10) - 1;
    const day = dateParts[1];
    const year = dateParts[2];
    const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
    const monthAbbr = months[monthIndex];

    const timeParts = timeCounted.split(':');
    let hours = parseInt(timeParts[0], 10);
    const minutes = timeParts[1];
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12; // The hour '0' should be '12'
    const formattedHours = String(hours).padStart(2, '0');

    let shiftSuffix = "";
    if (shiftType === "Day Shift") {
        shiftSuffix = "DS";
    } else if (shiftType === "Night Shift") {
        shiftSuffix = "NS";
    }

    const filename = `PMC_${monthAbbr}_${day}_${year}_${formattedHours}${minutes}_${ampm}_${shiftSuffix}.xlsx`;


    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

    await workbook.xlsx.write(res); // Isulat ang workbook direktang sa response
    res.end(); // Tapusin ang response
});

// Simulan ang server
app.listen(port, () => {
    console.log(`Server listening at http://localhost:${port}`);
});