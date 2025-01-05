const express = require('express')
const bodyParser = require('body-parser')
const path = require('path')
const {
    getTemplates,
    getPlaceholders,
    generatePreview,
    fillTemplate,
} = require('./controllers')
const { TEMPLATE_DIR } = require('./config')

const app = express()
const PORT = 3000

app.use(bodyParser.json())
app.use(express.static(path.join(__dirname, 'public')))

// Routes
app.get('/get-templates', getTemplates)
app.get('/get-placeholders', getPlaceholders)
app.post('/preview-template', generatePreview)
app.post('/fill-template', fillTemplate)

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`)
})
