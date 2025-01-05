const express = require('express')
const path = require('path')
const fs = require('fs')
const xlsx = require('xlsx')
const app = express()
const port = 3000

const templatesDir = path.join(__dirname, 'templates')
const updatedFilesDir = path.join(__dirname, 'updated-files')

// Middleware to handle JSON requests
app.use(express.json())

// Serve static files like CSS/JS
app.use(express.static(path.join(__dirname, 'public')))

// Endpoint to fetch available templates
app.get('/get-templates', (req, res) => {
    const files = fs.readdirSync(templatesDir)
    const templates = files.filter((file) => file.endsWith('.xlsx'))
    res.json({ templates })
})

// Endpoint to get placeholders from the selected template
app.get('/get-placeholders', (req, res) => {
    const fileName = req.query.file
    const filePath = path.join(templatesDir, fileName)

    const workbook = xlsx.readFile(filePath)
    const sheet = workbook.Sheets[workbook.SheetNames[0]]

    const cellValues = Object.values(sheet).map((cell) => cell.v)
    const placeholderPattern = /{{(.*?)}}/g
    const placeholders = new Set()

    cellValues.forEach((cellValue) => {
        let match
        while ((match = placeholderPattern.exec(cellValue)) !== null) {
            placeholders.add(match[1])
        }
    })

    res.json({ placeholders: Array.from(placeholders) })
})

// Endpoint to generate preview Excel file with user input (without saving)
app.post('/preview-template', (req, res) => {
    const { fileName, formData } = req.body
    const filePath = path.join(templatesDir, fileName)

    // Read the Excel template
    const workbook = xlsx.readFile(filePath)
    const sheet = workbook.Sheets[workbook.SheetNames[0]]

    // Replace placeholders with user input
    for (const [placeholder, value] of Object.entries(formData)) {
        const regex = new RegExp(`{{${placeholder}}}`, 'g')
        Object.keys(sheet).forEach((cellKey) => {
            const cell = sheet[cellKey]
            if (typeof cell.v === 'string') {
                cell.v = cell.v.replace(regex, value)
            }
        })
    }

    // Create a new workbook with updated data
    const updatedWorkbook = xlsx.utils.book_new()
    xlsx.utils.book_append_sheet(updatedWorkbook, sheet, workbook.SheetNames[0])

    // Send the updated workbook as a preview (buffer)
    const buffer = xlsx.write(updatedWorkbook, {
        bookType: 'xlsx',
        type: 'buffer',
    })
    res.send(buffer)
})

// Endpoint to generate the final filled Excel file (for download)
app.post('/fill-template', (req, res) => {
    const { fileName, formData } = req.body
    const filePath = path.join(templatesDir, fileName)

    // Read the Excel template
    const workbook = xlsx.readFile(filePath)
    const sheet = workbook.Sheets[workbook.SheetNames[0]]

    // Replace placeholders with user input
    for (const [placeholder, value] of Object.entries(formData)) {
        const regex = new RegExp(`{{${placeholder}}}`, 'g')
        Object.keys(sheet).forEach((cellKey) => {
            const cell = sheet[cellKey]
            if (typeof cell.v === 'string') {
                cell.v = cell.v.replace(regex, value)
            }
        })
    }

    // Create a new workbook with updated data
    const updatedWorkbook = xlsx.utils.book_new()
    xlsx.utils.book_append_sheet(updatedWorkbook, sheet, workbook.SheetNames[0])

    // Send the updated file for download
    const buffer = xlsx.write(updatedWorkbook, {
        bookType: 'xlsx',
        type: 'buffer',
    })
    res.setHeader('Content-Disposition', `attachment; filename=${fileName}`)
    res.send(buffer)
})

// Start the server
app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`)
})
