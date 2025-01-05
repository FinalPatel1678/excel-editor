const express = require('express')
const bodyParser = require('body-parser')
const multer = require('multer')
const fs = require('fs')
const path = require('path')
const ExcelJS = require('exceljs')

const app = express()
const PORT = 3000

app.use(bodyParser.json())
app.use(express.static(path.join(__dirname, 'public')))

const TEMPLATE_DIR = path.join(__dirname, 'templates')

// Endpoint to get the list of templates
app.get('/get-templates', (req, res) => {
    fs.readdir(TEMPLATE_DIR, (err, files) => {
        if (err) {
            console.error('Error reading template directory:', err)
            return res.status(500).json({ error: 'Unable to load templates' })
        }

        const excelFiles = files.filter((file) => file.endsWith('.xlsx'))
        res.json({ templates: excelFiles })
    })
})

// Endpoint to extract placeholders from a selected template
app.get('/get-placeholders', async (req, res) => {
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
})

// Endpoint to generate a preview of the Excel sheet
app.post('/preview-template', async (req, res) => {
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
        const sheet = workbook.worksheets[0] // Assume first sheet for preview
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
})

// Endpoint to fill the template and return the filled file
app.post('/fill-template', async (req, res) => {
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

        // Save the filled workbook to a buffer
        const buffer = await workbook.xlsx.writeBuffer()

        // Set proper headers and send the file for download
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
})

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`)
})
