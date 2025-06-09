// my-excel-app/api/generate-excel.js
const ExcelJS = require('exceljs'); // Changed from XLSX to ExcelJS

module.exports = async (req, res) => {
    if (req.method !== 'POST') {
        return res.status(405).send('Method Not Allowed');
    }

    try {
        const { inventoryData, dateShift, timeCounted, shiftType } = req.body;

        if (!inventoryData) {
            return res.status(400).send('Inventory data is required.');
        }

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Inventory');

        // --- Header Information ---
        // Title
        worksheet.mergeCells('A1:D1');
        worksheet.getCell('A1').value = 'PACKAGING MATERIALS DAILY COUNT SHEET - COHIN (BLDG. 3&6)';
        worksheet.getCell('A1').font = { bold: true, size: 14 };
        worksheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };

        // Date / Shift
        worksheet.getCell('A3').value = 'Date / Shift :';
        worksheet.getCell('B3').value = dateShift;
        worksheet.getCell('B3').font = { italic: true, bold: true };
        worksheet.getCell('D3').value = 'Shift:';
        worksheet.getCell('E3').value = shiftType; // Note: This will go beyond D column, consider merging if needed
        worksheet.getCell('E3').font = { italic: true, bold: true };

        // Time Counted
        worksheet.getCell('A4').value = 'TIME COUNTED:';
        worksheet.getCell('B4').value = timeCounted;
        worksheet.getCell('B4').font = { italic: true, bold: true };

        // Empty row
        worksheet.addRow([]); // Row 5 is empty

        // --- Main Headers (Row 6) ---
        const headers = ["ITEM CODE", "ACTUAL COUNT BREAKDOWN", "TOTAL", "REMARKS"];
        const headerRow = worksheet.getRow(6);
        headerRow.values = headers;
        headerRow.eachCell({ includeEmpty: false }, (cell) => {
            cell.font = { bold: true };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } }; // Tailwind's gray-200
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });
        worksheet.getRow(6).height = 20; // Set header row height

        // --- Column Widths ---
        worksheet.columns = [
            { header: 'ITEM CODE', key: 'code', width: 25 },
            { header: 'ACTUAL COUNT BREAKDOWN', key: 'countBreakdown', width: 40 },
            { header: 'TOTAL', key: 'total', width: 15 },
            { header: 'REMARKS', key: 'remarks', width: 30 }
        ];

        // --- Data Rows (Starting from Row 7) ---
        // Helper functions from your HTML (adapt for Node.js)
        const calculateTotal = (countArr) => {
            if (!Array.isArray(countArr)) return 0;
            let total = 0;
            countArr.forEach(part => {
                const trimmedVal = part.trim();
                if (trimmedVal === "") return;
                try {
                    // Use a safer eval or a proper math parser if expressions are complex
                    const evaluated = eval(trimmedVal.replace(/×/g, '*')); // Replace '×' with '*' for eval
                    if (!isNaN(evaluated)) {
                        total += evaluated;
                    }
                } catch (e) {
                    // Handle invalid expressions
                    console.error("Error evaluating count part:", trimmedVal, e);
                }
            });
            return total;
        };

        const generateExcelFormula = (countArr) => {
            if (!Array.isArray(countArr) || countArr.length === 0) {
                return "0";
            }
            const parsedParts = countArr.map(part => {
                const trimmedVal = part.trim();
                if (trimmedVal === "") return "0";
                let formattedPart = trimmedVal.replace(/×/g, '*');
                if (formattedPart.includes('+') || formattedPart.includes('-')) {
                    return `(${formattedPart})`;
                }
                return formattedPart;
            });
            return `=${parsedParts.join('+')}`;
        };

        // Function to get color based on remark (adapting from Tailwind colors)
        const getRemarkBgColor = (remark) => {
            const r = remark.toLowerCase();
            if (r.includes("hold")) return 'FFFDE68A'; // yellow-300
            if (r.includes("approved")) return 'FFFBCFE8'; // pink-300
            if (r.includes("first out") || r.includes("old")) return 'FFA7F3D0'; // green-300
            return 'FFFFFFFF'; // white
        };

        inventoryData.forEach(item => {
            const countBreakdownString = item.count.join(' | ');
            const remarksString = item.remarks.join(' | ');
            const totalNumerical = calculateTotal(item.count);
            const totalFormula = generateExcelFormula(item.count);

            const row = worksheet.addRow([item.code, countBreakdownString, totalFormula, remarksString]);

            // Apply borders to all cells in the row
            row.eachCell({ includeEmpty: false }, (cell) => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });

            // Specific styling for cells
            // Item Code
            row.getCell(1).font = { italic: true, bold: true }; // Make item code italic and bold

            // Total - set as formula and bold
            row.getCell(3).value = { formula: totalFormula.substring(1), result: totalNumerical }; // Set formula, remove '='
            row.getCell(3).font = { bold: true };

            // Remarks cell background color based on actual remarks (can be multiple)
            const remarksCell = row.getCell(4);
            const remarksArray = Array.isArray(item.remarks) ? item.remarks : [];
            let appliedColor = 'FFFFFFFF'; // Default to white
            for (const r of remarksArray) {
                const color = getRemarkBgColor(r);
                if (color !== 'FFFFFFFF') { // If any remark has a specific color, use it
                    appliedColor = color;
                    break; // Use the first specific color found
                }
            }
            remarksCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: appliedColor } };

        });

        // Set autoFilter to all columns containing data
        worksheet.autoFilter = {
            from: 'A6',
            to: `D${worksheet.rowCount}` // Adjust to your actual data range
        };


        // Generate buffer for the Excel file
        const buffer = await workbook.xlsx.writeBuffer();

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="Styled_Countsheet_${dateShift.replace(/\//g, '-')}.xlsx"`);
        res.status(200).send(buffer);

    } catch (error) {
        console.error("Error generating Excel file:", error);
        res.status(500).send(`Failed to generate Excel file: ${error.message}`);
    }
};