const fs = require('fs')
const path = require('path')
const { TEMPLATE_DIR } = require('../config')

module.exports = async (req, res) => {
    try {
        const files = await fs.promises.readdir(TEMPLATE_DIR)
        const excelFiles = files.filter((file) => file.endsWith('.xlsx'))
        res.json({ templates: excelFiles })
    } catch (err) {
        console.error('Error reading template directory:', err)
        res.status(500).json({ error: 'Unable to load templates' })
    }
}
