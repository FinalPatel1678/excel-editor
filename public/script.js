const messages = {
    en: {
        templateLabel: 'Select Template:',
        instructionText:
            'Please select a template and fill in the required fields. After that, you can preview the template and download it with your filled data.',
        previewBtn: 'Get Preview',
        downloadBtn: 'Download File',
        downloadPdfBtn: 'Download PDF',
        loading: 'Loading templates...',
        error: 'Error fetching templates.',
        formError: 'Please fill all fields before downloading.',
        dynamicFieldsEmpty: 'No placeholders available for this template.',
        dynamicFieldsNoneFound: 'No placeholders found in the template.',
        warningEmptyField: '⚠️ Please fill this field.', // Tooltip for empty fields
        excelPreviewTitle: 'Template Preview (Editable)',
        languageSelectLabel: 'Select Language:',
    },
    hi: {
        templateLabel: 'टेम्पलेट चुनें:',
        instructionText:
            'कृपया एक टेम्पलट चुनें और आवश्यक फ़ील्ड भरें। उसके बाद, आप टेम्पलट का पूर्वावलोकन कर सकते हैं और भरे गए डेटा के साथ इसे डाउनलोड कर सकते हैं।',
        previewBtn: 'पूर्वावलोकन प्राप्त करें',
        downloadBtn: 'फ़ाइल डाउनलोड करें',
        downloadPdfBtn: 'पीडीएफ डाउनलोड करें',
        loading: 'टेम्पलट लोड हो रहा है...',
        error: 'टेम्पलट लोड करने में त्रुटि।',
        formError: 'डाउनलोड करने से पहले सभी फ़ील्ड भरें।',
        dynamicFieldsEmpty: 'इस टेम्पलट के लिए कोई प्लेसहोल्डर उपलब्ध नहीं है।',
        dynamicFieldsNoneFound: 'टेम्पलट में कोई प्लेसहोल्डर नहीं मिला।',
        warningEmptyField: '⚠️ कृपया यह फ़ील्ड भरें।',
        excelPreviewTitle: 'टेम्पलट पूर्वावलोकन (संपादन योग्य)',
        languageSelectLabel: 'भाषा चुनें:',
    },
    gu: {
        templateLabel: 'ટેમ્પલેટ પસંદ કરો:',
        instructionText:
            'કૃપા કરી એક ટેમ્પલેટ પસંદ કરો અને જરૂરી ફીલ્ડ ભરો. ત્યારબાદ, તમે ટેમ્પલેટનું પૂર્વાવલોકન કરી શકો છો અને ભરેલા ડેટા સાથે તેને ડાઉનલોડ કરી શકો છો.',
        previewBtn: 'પૂર્વાવલોકન મેળવો',
        downloadBtn: 'ફાઈલ ડાઉનલોડ કરો',
        downloadPdfBtn: 'પીડીએફ ડાઉનલોડ કરો',
        loading: 'ટેમ્પલેટ લોડ થઈ રહ્યું છે...',
        error: 'ટેમ્પલેટ લોડ કરવામાં ભૂલ.',
        formError: 'ડાઉનલોડ કરતા પહેલા તમામ ફીલ્ડ ભરાવા જોઈએ.',
        dynamicFieldsEmpty: 'આ ટેમ્પલેટ માટે કોઈ પ્લેસહોલ્ડર ઉપલબ્ધ નથી.',
        dynamicFieldsNoneFound: 'ટેમ્પલેટમાં કોઈ પ્લેસહોલ્ડર મળ્યા નથી.',
        warningEmptyField: '⚠️ કૃપા કરી આ ફીલ્ડ ભરો.',
        excelPreviewTitle: 'ટેમ્પલેટ પૂર્વાવલોકન (ફેરફાર યોગ્ય)',
        languageSelectLabel: 'ભાષા પસંદ કરો:',
    },
}

document.addEventListener('DOMContentLoaded', () => {
    const langSelect = document.getElementById('langSelect')
    const templateSelect = document.getElementById('template')
    const dynamicFields = document.getElementById('dynamicFields')
    const previewBtn = document.getElementById('previewBtn')
    const downloadBtn = document.getElementById('downloadBtn')
    const downloadPdfBtn = document.getElementById('downloadPdfBtn')
    const excelPreview = document.getElementById('excelPreview')
    const formError = document.getElementById('formError')
    const templateLoading = document.getElementById('templateLoading')
    const instructionText = document.getElementById('instructionText')
    const languageSelectLabel = document.getElementById('languageSelectLabel') // Added if needed

    let handsontableInstance

    // Common function to toggle loading state for buttons
    function toggleButtonLoadingState(
        button,
        isLoading,
        loadingText = messages[langSelect.value]?.loading || 'Loading...'
    ) {
        if (isLoading) {
            button.disabled = true
            button.dataset.originalText = button.textContent // Save original text
            button.textContent = loadingText // Set loading text
        } else {
            button.disabled = false
            button.textContent =
                button.dataset.originalText || button.textContent // Restore original text
        }
    }

    // Update UI text based on selected language
    function updateUI(language) {
        document.getElementById('templateLabel').textContent =
            messages[language].templateLabel
        instructionText.textContent = messages[language].instructionText
        previewBtn.textContent = messages[language].previewBtn
        downloadBtn.textContent = messages[language].downloadBtn
        downloadPdfBtn.textContent = messages[language].downloadPdfBtn
        templateLoading.textContent = messages[language].loading
        formError.textContent = messages[language].formError
        excelPreview.setAttribute(
            'aria-label',
            messages[language].excelPreviewTitle
        ) // Accessible title for preview
        if (languageSelectLabel) {
            languageSelectLabel.textContent =
                messages[language].languageSelectLabel
        }
    }

    // Highlight empty fields with localized tooltips
    function highlightEmptyFields(formData) {
        dynamicFields.querySelectorAll('input').forEach((input) => {
            if (!formData[input.name]) {
                input.style.border = '1px solid red' // Highlight empty field
                input.title = messages[langSelect.value].warningEmptyField // Tooltip
            } else {
                input.style.border = '' // Remove highlight when filled
                input.title = '' // Remove tooltip
            }
        })
    }

    // Fetch available templates
    function fetchTemplates() {
        templateLoading.textContent = messages[langSelect.value].loading
        fetch('api/get-templates')
            .then((res) => res.json())
            .then((data) => {
                populateTemplateSelect(data.templates)
            })
            .catch((error) => {
                console.error(messages[langSelect.value].error, error)
                templateLoading.textContent = messages[langSelect.value].error
            })
    }

    function populateTemplateSelect(templates) {
        templateSelect.innerHTML = ''
        templates.forEach((template) => {
            const option = document.createElement('option')
            option.value = template
            option.textContent = template
            templateSelect.appendChild(option)
        })
        templateLoading.textContent = ''

        // Automatically generate preview for the selected template
        if (templateSelect.value) {
            fetchPlaceholders(templateSelect.value)
            generatePreview()
        }
    }

    // Fetch placeholders for the selected template
    function fetchPlaceholders(template) {
        dynamicFields.innerHTML = ''
        excelPreview.innerHTML = '' // Clear preview

        fetch(`api/get-placeholders?file=${template}`)
            .then((res) => res.json())
            .then((data) => populateDynamicFields(data.placeholders_by_sheet))
            .catch((error) => {
                console.error('Error fetching placeholders:', error)
                dynamicFields.innerHTML = `<p>${
                    messages[langSelect.value].dynamicFieldsEmpty
                }</p>`
            })
    }

    // Populate dynamic fields
    function populateDynamicFields(placeholdersBySheet) {
        dynamicFields.innerHTML = '' // Clear existing fields

        // Filter sheets that have placeholders
        const sheetsWithPlaceholders = Object.entries(
            placeholdersBySheet
        ).filter(([_, placeholders]) => placeholders.length > 0)

        // If no sheets have placeholders, show a message
        if (sheetsWithPlaceholders.length === 0) {
            dynamicFields.innerHTML = `<p>${
                messages[langSelect.value].dynamicFieldsNoneFound
            }</p>`
            return
        }

        // Iterate through sheets with placeholders and add sections
        sheetsWithPlaceholders.forEach(([sheetName, placeholders]) => {
            const section = document.createElement('div')
            section.style.marginBottom = '20px'
            section.style.border = '1px solid #ccc'
            section.style.borderRadius = '5px'
            section.style.padding = '10px'
            section.style.backgroundColor = '#f9f9f9'

            // Add a heading for the sheet
            const heading = document.createElement('h3')
            heading.textContent = `Sheet: ${sheetName}`
            heading.style.marginBottom = '10px'
            heading.style.color = '#007bff'
            heading.style.fontSize = '1.2em'
            section.appendChild(heading)

            // Add placeholders for the sheet
            placeholders.forEach((placeholder) => {
                const label = document.createElement('label')
                label.textContent = `${placeholder}: `
                label.style.display = 'block'
                label.style.marginBottom = '5px'
                label.style.fontWeight = 'bold'

                const input = document.createElement('input')
                input.name = placeholder
                input.placeholder = `${messages[
                    langSelect.value
                ].warningEmptyField.replace('⚠️', '')} ${placeholder}`
                input.style.width = '100%'
                input.style.marginBottom = '10px'
                input.style.padding = '5px'
                input.style.border = '1px solid #ccc'
                input.style.borderRadius = '3px'

                section.appendChild(label)
                section.appendChild(input)
            })

            dynamicFields.appendChild(section)
        })
    }

    // Highlight empty fields with red borders and add warning icon
    function highlightEmptyFields(formData) {
        dynamicFields.querySelectorAll('input').forEach((input) => {
            if (!formData[input.name]) {
                input.style.border = '1px solid red' // Highlight empty field
                addWarningIcon(input) // Add warning icon for empty fields
            } else {
                input.style.border = '' // Remove highlight when filled
                removeWarningIcon(input) // Remove warning icon
            }
        })
    }

    // Add warning icon next to empty fields
    function addWarningIcon(input) {
        if (!input.nextElementSibling) {
            const warningIcon = document.createElement('span')
            warningIcon.textContent = '⚠️'
            warningIcon.style.color = 'red'
            warningIcon.style.marginLeft = '5px'
            input.parentNode.appendChild(warningIcon)
        }
    }

    // Remove warning icon
    function removeWarningIcon(input) {
        if (
            input.nextElementSibling &&
            input.nextElementSibling.textContent === '⚠️'
        ) {
            input.parentNode.removeChild(input.nextElementSibling)
        }
    }

    // Handle template selection change
    templateSelect.addEventListener('change', () => {
        fetchPlaceholders(templateSelect.value)
        generatePreview() // Generate preview whenever template is changed
    })

    // Handle preview button (Allow preview without filling data)
    previewBtn.addEventListener('click', () => {
        formError.textContent = '' // Clear any previous error
        toggleButtonLoadingState(previewBtn, true, 'Generating Preview...')
        generatePreview().finally(() =>
            toggleButtonLoadingState(previewBtn, false)
        ) // Re-enable button
    })

    // Generate preview
    function generatePreview() {
        const formData = Array.from(
            dynamicFields.querySelectorAll('input')
        ).reduce((data, input) => {
            data[input.name] = input.value
            return data
        }, {})

        return fetch('api/preview-template', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ fileName: templateSelect.value, formData }),
        })
            .then((res) => res.json())
            .then((excelData) => {
                const container = document.createElement('div')
                const tabsContainer = document.createElement('div')
                const contentContainer = document.createElement('div')

                container.appendChild(tabsContainer)
                container.appendChild(contentContainer)
                excelPreview.innerHTML = ''
                excelPreview.appendChild(container)

                tabsContainer.style.marginBottom = '10px'
                tabsContainer.style.display = 'flex'
                tabsContainer.style.flexWrap = 'wrap'

                const instances = {}

                Object.keys(excelData).forEach((sheetName, index) => {
                    const tab = document.createElement('button')
                    tab.textContent = sheetName
                    tab.style.margin = '0 5px'
                    tab.style.padding = '5px 10px'
                    tab.style.cursor = 'pointer'

                    const sheetDiv = document.createElement('div')
                    sheetDiv.style.display = index === 0 ? 'block' : 'none'
                    sheetDiv.style.width = '100%'
                    sheetDiv.style.height = '400px'
                    contentContainer.appendChild(sheetDiv)

                    instances[sheetName] = new Handsontable(sheetDiv, {
                        data: excelData[sheetName].map(
                            (row) => row.map((cellData) => cellData.value || '') // Map to just the value for the data
                        ),
                        colHeaders:
                            excelData[sheetName].length > 0
                                ? Object.keys(excelData[sheetName][0])
                                : [], // Dynamically set column headers
                        rowHeaders: true,
                        height: '400px',
                        licenseKey: 'non-commercial-and-evaluation',
                        columnSorting: true,
                        cells: (row, col) => {
                            const cellData = excelData[sheetName][row][col]
                            const cellStyle = cellData.styles || {} // Get the cell styles

                            return {
                                renderer: (
                                    instance,
                                    td,
                                    row,
                                    col,
                                    prop,
                                    value,
                                    cellProperties
                                ) => {
                                    // Apply default text renderer
                                    Handsontable.renderers.TextRenderer(
                                        instance,
                                        td,
                                        row,
                                        col,
                                        prop,
                                        value,
                                        cellProperties
                                    )

                                    // Apply background color if available
                                    if (cellStyle.bg_color) {
                                        td.style.backgroundColor =
                                            cellStyle.bg_color
                                    }

                                    // Apply font styles if available
                                    if (cellStyle.font) {
                                        td.style.color =
                                            cellStyle.font.color || ''
                                        td.style.fontSize =
                                            cellStyle.font.size || ''
                                        td.style.fontWeight = cellStyle.font
                                            .bold
                                            ? 'bold'
                                            : 'normal'
                                        td.style.fontStyle = cellStyle.font
                                            .italic
                                            ? 'italic'
                                            : 'normal'
                                        td.style.textDecoration = cellStyle.font
                                            .underline
                                            ? 'underline'
                                            : 'none'
                                    }

                                    // Apply borders if available
                                    if (cellStyle.border) {
                                        td.style.borderTop =
                                            cellStyle.border.top.style || ''
                                        td.style.borderRight =
                                            cellStyle.border.right.style || ''
                                        td.style.borderBottom =
                                            cellStyle.border.bottom.style || ''
                                        td.style.borderLeft =
                                            cellStyle.border.left.style || ''

                                        // Apply border color
                                        td.style.borderTopColor =
                                            cellStyle.border.top.color || ''
                                        td.style.borderRightColor =
                                            cellStyle.border.right.color || ''
                                        td.style.borderBottomColor =
                                            cellStyle.border.bottom.color || ''
                                        td.style.borderLeftColor =
                                            cellStyle.border.left.color || ''
                                    }

                                    // Apply number format (if any)
                                    if (cellStyle.number_format) {
                                        td.style.format =
                                            cellStyle.number_format
                                    }

                                    // Apply hyperlink (if any)
                                    if (cellStyle.hyperlink) {
                                        td.innerHTML = `<a href="${cellStyle.hyperlink}" target="_blank">${td.innerHTML}</a>`
                                    }

                                    // Apply row height (if available)
                                    // if (cellStyle.row_height) {
                                    //     sheetDiv.style.height =
                                    //         cellStyle.row_height + 'px'
                                    // }

                                    // Apply merged cells (if available)
                                    if (cellStyle.merged) {
                                        td.style.border = '2px solid #000' // Example for merged cells
                                    }

                                    // Add comments (if available)
                                    if (cellStyle.comment) {
                                        td.title = cellStyle.comment
                                    }

                                    // Handle formula if available
                                    if (cellStyle.formula) {
                                        td.style.fontStyle = 'italic'
                                        td.style.color = 'blue'
                                        td.title = `Formula: ${cellStyle.formula}`
                                    }
                                },
                            }
                        },
                    })

                    tab.addEventListener('click', () => {
                        Object.values(instances).forEach((instance) => {
                            instance.rootElement.style.display = 'none'
                        })
                        sheetDiv.style.display = 'block'
                        instances[sheetName].render()

                        tabsContainer
                            .querySelectorAll('button')
                            .forEach((btn) => {
                                btn.style.backgroundColor = ''
                                btn.style.color = ''
                            })
                        tab.style.backgroundColor = '#007bff'
                        tab.style.color = 'white'
                    })

                    if (index === 0) {
                        tab.style.backgroundColor = '#007bff'
                        tab.style.color = 'white'
                    }

                    tabsContainer.appendChild(tab)
                })

                if (handsontableInstance) {
                    Object.values(handsontableInstance).forEach((instance) =>
                        instance.destroy()
                    )
                }
                handsontableInstance = instances
            })
            .catch((error) => {
                console.error('Error generating preview:', error)
                alert('Failed to generate preview. Please try again later.')
            })
    }

    // Handle download button
    downloadBtn.addEventListener('click', () => {
        const formData = Array.from(
            dynamicFields.querySelectorAll('input')
        ).reduce((data, input) => {
            data[input.name] = input.value
            return data
        }, {})

        // Check if any fields are empty
        const emptyFields = Object.values(formData).some(
            (value) => value === ''
        )

        if (emptyFields) {
            formError.textContent = messages[langSelect.value].formError
            highlightEmptyFields(formData) // Add visual feedback
        } else {
            toggleButtonLoadingState(downloadBtn, true, 'Downloading...')
            fetch('api/fill-template', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    fileName: templateSelect.value,
                    formData,
                }),
            })
                .then((res) => res.blob())
                .then((blob) => {
                    const url = window.URL.createObjectURL(blob)
                    const a = document.createElement('a')
                    a.href = url
                    a.download = `${templateSelect.value.replace(
                        '.xlsx',
                        '_filled.xlsx'
                    )}`
                    document.body.appendChild(a)
                    a.click()
                    a.remove()
                })
                .catch((error) => {
                    console.error('Error downloading file:', error)
                })
                .finally(() => toggleButtonLoadingState(downloadBtn, false)) // Re-enable button
        }
    })

    // Handle PDF download button
    downloadPdfBtn.addEventListener('click', () => {
        const formData = Array.from(
            dynamicFields.querySelectorAll('input')
        ).reduce((data, input) => {
            data[input.name] = input.value
            return data
        }, {})

        // Check if any fields are empty
        const emptyFields = Object.values(formData).some(
            (value) => value === ''
        )

        if (emptyFields) {
            formError.textContent = messages[langSelect.value].formError
            highlightEmptyFields(formData)
        } else {
            toggleButtonLoadingState(downloadPdfBtn, true, 'Downloading PDF...')
            fetch('api/fill-template-pdf', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    fileName: templateSelect.value,
                    formData,
                }),
            })
                .then((res) => res.blob())
                .then((blob) => {
                    const url = window.URL.createObjectURL(blob)
                    const a = document.createElement('a')
                    a.href = url
                    a.download = `${templateSelect.value.replace(
                        '.xlsx',
                        '_filled.pdf'
                    )}`
                    document.body.appendChild(a)
                    a.click()
                    a.remove()
                })
                .catch((error) => {
                    console.error('Error downloading PDF:', error)
                })
                .finally(() => toggleButtonLoadingState(downloadPdfBtn, false)) // Re-enable button
        }
    })

    // Change language when user selects different option
    langSelect.addEventListener('change', () => {
        updateUI(langSelect.value)
        templateLoading.textContent = '' // Fix loading label issue
    })

    // Set the default language to Gujarati and update UI
    langSelect.value = 'en'
    updateUI('en')

    // Fetch templates initially
    fetchTemplates()
})
