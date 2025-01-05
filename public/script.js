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
            .then((data) => populateDynamicFields(data.placeholders))
            .catch((error) => {
                console.error('Error fetching placeholders:', error)
                dynamicFields.innerHTML = `<p>${
                    messages[langSelect.value].dynamicFieldsEmpty
                }</p>`
            })
    }

    // Populate dynamic fields
    function populateDynamicFields(placeholders) {
        if (placeholders.length === 0) {
            dynamicFields.innerHTML = `<p>${
                messages[langSelect.value].dynamicFieldsNoneFound
            }</p>`
            return
        }
        placeholders.forEach((placeholder) => {
            const label = document.createElement('label')
            label.textContent = `${placeholder}: `
            const input = document.createElement('input')
            input.name = placeholder
            input.placeholder = `${messages[
                langSelect.value
            ].warningEmptyField.replace('⚠️', '')} ${placeholder}`
            dynamicFields.appendChild(label)
            dynamicFields.appendChild(input)
            dynamicFields.appendChild(document.createElement('br'))
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
                if (handsontableInstance) handsontableInstance.destroy()
                handsontableInstance = new Handsontable(excelPreview, {
                    data: excelData,
                    colHeaders: true,
                    rowHeaders: true,
                    height: '400px', // Set a max height for the preview container
                    licenseKey: 'non-commercial-and-evaluation',
                    columnSorting: true, // Sorting can help if the table is large
                })
            })
            .catch((error) => {
                console.error('Error generating preview:', error)
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
