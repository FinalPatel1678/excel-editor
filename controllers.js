const fs = require('fs')
const path = require('path')
const ExcelJS = require('exceljs')
const { TEMPLATE_DIR } = require('./config')

// Endpoint to get the list of templates
async function getTemplates(req, res) {
    try {
        const files = await fs.promises.readdir(TEMPLATE_DIR)
        const excelFiles = files.filter((file) => file.endsWith('.xlsx'))
        res.json({ templates: excelFiles })
    } catch (err) {
        console.error('Error reading template directory:', err)
        res.status(500).json({ error: 'Unable to load templates' })
    }
}

// Endpoint to extract placeholders from a selected template
async function getPlaceholders(req, res) {
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

// Endpoint to generate a preview of the Excel sheet
async function generatePreview(req, res) {
    const { fileName, formData } = req.body
    if (!fileName || !formData) {
        return res
            .status(400)
            .json({ error: 'File name and form data are required' })
    }

    const filePath = path.join(TEMPLATE_DIR, fileName)
    try {
        const workbook = new ExcelJS.Workbook()
        await workbook.xlsx.readFile(filePath)

        workbook.eachSheet((sheet) => {
            sheet.eachRow((row) => {
                row.eachCell((cell) => {
                    if (
                        typeof cell.value === 'string' &&
                        cell.value.includes('{{') &&
                        cell.value.includes('}}')
                    ) {
                        const placeholder = cell.value
                            .match(/{{(.*?)}}/)[1]
                            .trim()
                        if (formData[placeholder]) {
                            cell.value = formData[placeholder]
                        }
                    }
                })
            })
        })

        // Convert the workbook to JSON data for preview
        const previewData = []
        const sheet = workbook.worksheets[0]
        sheet.eachRow((row) => {
            const rowData = []
            row.eachCell((cell) => {
                rowData.push(cell.value || '')
            })
            previewData.push(rowData)
        })

        res.json(previewData)
    } catch (error) {
        console.error('Error generating preview:', error)
        res.status(500).json({ error: 'Failed to generate preview' })
    }
}

// Endpoint to fill the template and return the filled file
async function fillTemplate(req, res) {
    const { fileName, formData } = req.body
    if (!fileName || !formData) {
        return res
            .status(400)
            .json({ error: 'File name and form data are required' })
    }

    const filePath = path.join(TEMPLATE_DIR, fileName)
    try {
        const workbook = new ExcelJS.Workbook()
        await workbook.xlsx.readFile(filePath)

        workbook.eachSheet((sheet) => {
            sheet.eachRow((row) => {
                row.eachCell((cell) => {
                    if (
                        typeof cell.value === 'string' &&
                        cell.value.includes('{{') &&
                        cell.value.includes('}}')
                    ) {
                        const placeholder = cell.value
                            .match(/{{(.*?)}}/)[1]
                            .trim()
                        if (formData[placeholder]) {
                            cell.value = formData[placeholder]
                        }
                    }
                })
            })
        })

        const buffer = await workbook.xlsx.writeBuffer()
        res.setHeader(
            'Content-Disposition',
            `attachment; filename=filled_${fileName}`
        )
        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        res.send(buffer)
    } catch (error) {
        console.error('Error filling template:', error)
        res.status(500).json({ error: 'Failed to fill the template' })
    }
}

module.exports = {
    getTemplates,
    getPlaceholders,
    generatePreview,
    fillTemplate,
}
