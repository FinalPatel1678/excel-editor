const fs = require('fs')
const path = require('path')
const ExcelJS = require('exceljs')
const { TEMPLATE_DIR } = require('../config')

module.exports = async (req, res) => {
    const fileName = req.query.file
    if (!fileName) {
        return res.status(400).json({ error: 'Template file name is required' })
    }

    const filePath = path.join(TEMPLATE_DIR, fileName)
    try {
        const workbook = new ExcelJS.Workbook()
        await workbook.xlsx.readFile(filePath)

        const placeholders = new Set()
        workbook.eachSheet((sheet) => {
            sheet.eachRow((row) => {
                row.eachCell((cell) => {
                    const value = cell.value
                    if (
                        typeof value === 'string' &&
                        value.includes('{{') &&
                        value.includes('}}')
                    ) {
                        const matches = value.match(/{{(.*?)}}/g)
                        matches.forEach((match) => {
                            placeholders.add(match.replace(/{{|}}/g, '').trim())
                        })
                    }
                })
            })
        })
        res.json({ placeholders: Array.from(placeholders) })
    } catch (error) {
        console.error('Error extracting placeholders:', error)
        res.status(500).json({
            error: 'Failed to extract placeholders from the template',
        })
    }
}
