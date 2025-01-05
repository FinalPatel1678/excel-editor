const ExcelJS = require('exceljs')
const path = require('path')
const PDFDocument = require('pdfkit')
const { TEMPLATE_DIR } = require('../config')

module.exports = async (req, res) => {
    try {
        const { fileName, formData } = req.body

        // Load Excel file
        const workbook = new ExcelJS.Workbook()
        await workbook.xlsx.readFile(path.join(TEMPLATE_DIR, fileName))

        // Fill placeholders in Excel
        const worksheet = workbook.getWorksheet(1)
        worksheet.eachRow((row) => {
            row.eachCell((cell) => {
                if (typeof cell.value === 'string') {
                    Object.keys(formData).forEach((placeholder) => {
                        if (cell.value.includes(placeholder)) {
                            cell.value = cell.value.replace(
                                `{${placeholder}}`,
                                formData[placeholder]
                            )
                        }
                    })
                }
            })
        })

        // Create a PDF document
        const pdfDoc = new PDFDocument()
        const chunks = []
        pdfDoc.on('data', (chunk) => chunks.push(chunk))
        pdfDoc.on('end', () => {
            const pdfBuffer = Buffer.concat(chunks)
            res.setHeader('Content-Type', 'application/pdf')
            res.setHeader(
                'Content-Disposition',
                `attachment; filename="${fileName.replace('.xlsx', '.pdf')}"`
            )
            res.send(pdfBuffer)
        })

        // Write Excel data to the PDF
        worksheet.eachRow((row, rowIndex) => {
            row.eachCell((cell, colIndex) => {
                const fontSize = cell.font?.size || 12
                const align = cell.alignment?.horizontal || 'left'
                const fontColor = cell.font?.color?.argb || 'black'
                const bgColor = cell.fill?.fgColor?.argb || 'white'

                // Set font, color, background, and alignment
                pdfDoc
                    .fontSize(fontSize)
                    .fillColor(fontColor)
                    .text(cell.value || '', {
                        width: 500,
                        align: align, // Text alignment (left, center, right)
                    })

                // Apply background color if needed (for cells with background fill)
                if (bgColor !== 'white') {
                    pdfDoc.rect(10, 10, 100, 20).fill(bgColor) // Example rectangle for background
                }
            })
            pdfDoc.moveDown()
        })

        pdfDoc.end()
    } catch (err) {
        console.error('Error generating PDF:', err)
        res.status(500).json({ error: 'Error generating PDF' })
    }
}
