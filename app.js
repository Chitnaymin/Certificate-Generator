const fs = require("fs");
const PDFDocument = require("pdfkit");
const ExcelJS = require('exceljs');

const workbook = new ExcelJS.Workbook();

// Load the Excel file
workbook.xlsx.readFile('Student List.xlsx')
    .then(() => {
        const worksheet = workbook.getWorksheet(1);

        worksheet.eachRow({ includeEmpty: true }, (row) => {
            const name = row.getCell(1).value;
            createCertificate(name);
        });

    })
    .then(() => {
        console.log('All certificates created successfully.');
    })
    .catch(error => {
        console.error('Error reading Excel file:', error.message);
    });

//createCertificate('Min Khant Kyaw');

function createCertificate(name) {
    // Finalize the PDF and end the stream
    // Create a new PDF document for each certificate
    const doc = new PDFDocument({
        layout: "landscape",
        size: "A4",
    });

    // Pipe the PDF into an name.pdf file
    const outputPath = `Export/${name}.pdf`;
    doc.pipe(fs.createWriteStream(outputPath));

    // Draw the certificate image
    doc.image("images/CIS Kahtein Volunteer Certificate of Participation.png", 0, 0, { width: 842 });

    // Remember to download the font
    // Set the font to Times New Roman
    doc.font("fonts/TIMESBD.ttf");
    doc.fillColor('#32312F');
    // Draw the name
    doc.fontSize(45).text(name, 75, 290, {
        align: "center",
    });

    doc.end();
}
