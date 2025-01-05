const ExcelJS = require('exceljs')
const path = require('path')
const { TEMPLATE_DIR } = require('../config')

module.exports = async (req, res) => {
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
