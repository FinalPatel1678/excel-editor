document.addEventListener('DOMContentLoaded', () => {
    const langSelect = document.getElementById('langSelect')
    const templateSelect = document.getElementById('template')
    const dynamicFields = document.getElementById('dynamicFields')
    const previewBtn = document.getElementById('previewBtn')
    const downloadBtn = document.getElementById('downloadBtn')
    const excelPreview = document.getElementById('excelPreview')
    const formError = document.getElementById('formError')
    const templateLoading = document.getElementById('templateLoading')
    const instructionText = document.getElementById('instructionText')

    let handsontableInstance // Handsontable instance

    const messages = {
        en: {
            templateLabel: 'Select Template:',
            instructionText:
                'Please select a template and fill in the required fields. After that, you can preview the template and download it with your filled data.',
            previewBtn: 'Get Preview',
            downloadBtn: 'Download File',
            loading: 'Loading templates...',
            error: 'Error fetching templates.',
            formError: 'Please fill all fields before downloading.',
        },
        hi: {
            templateLabel: 'टेम्पलेट चुनें:',
            instructionText:
                'कृपया एक टेम्पलट चुनें और आवश्यक फ़ील्ड भरें। उसके बाद, आप टेम्पलट का पूर्वावलोकन कर सकते हैं और भरे गए डेटा के साथ इसे डाउनलोड कर सकते हैं।',
            previewBtn: 'पूर्वावलोकन प्राप्त करें',
            downloadBtn: 'फ़ाइल डाउनलोड करें',
            loading: 'टेम्पलट लोड हो रहा है...',
            error: 'टेम्पलट लोड करने में त्रुटि।',
            formError: 'डाउनलोड करने से पहले सभी फ़ील्ड भरें।',
        },
        gu: {
            templateLabel: 'ટેમ્પલેટ પસંદ કરો:',
            instructionText:
                'કૃપા કરી એક ટેમ્પલેટ પસંદ કરો અને જરૂરી ફીલ્ડ ભરો. ત્યારબાદ, તમે ટેમ્પલેટનું પૂર્વાવલોકન કરી શકો છો અને ભરેલા ડેટા સાથે તેને ડાઉનલોડ કરી શકો છો.',
            previewBtn: 'પૂર્વાવલોકન મેળવો',
            downloadBtn: 'ફાઈલ ડાઉનલોડ કરો',
            loading: 'ટેમ્પલેટ લોડ થઈ રહ્યું છે...',
            error: 'ટેમ્પલેટ લોડ કરવામાં ભૂલ.',
            formError: 'ડાઉનલોડ કરતા પહેલા તમામ ફીલ્ડ ભરાવા જોઈએ.',
        },
    }

    // Update UI text based on selected language
    function updateUI(language) {
        document.getElementById('templateLabel').textContent =
            messages[language].templateLabel
        instructionText.textContent = messages[language].instructionText
        previewBtn.textContent = messages[language].previewBtn
        downloadBtn.textContent = messages[language].downloadBtn
        templateLoading.textContent = messages[language].loading
        formError.textContent = messages[language].formError
    }

    // Fetch available templates
    function fetchTemplates() {
        templateLoading.textContent = messages[langSelect.value].loading
        fetch('/get-templates')
            .then((res) => res.json())
            .then((data) => {
                populateTemplateSelect(data.templates)
            })
            .catch((error) => {
                console.error(messages[langSelect.value].error, error)
                templateLoading.textContent = 'Failed to load templates!'
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
        if (templateSelect.value) {
            fetchPlaceholders(templateSelect.value)
            generatePreview() // Automatically generate the preview for the default template
        }
    }

    // Fetch placeholders for the selected template
    function fetchPlaceholders(template) {
        dynamicFields.innerHTML = ''
        excelPreview.innerHTML = '' // Clear preview

        fetch(`/get-placeholders?file=${template}`)
            .then((res) => res.json())
            .then((data) => populateDynamicFields(data.placeholders))
            .catch((error) => {
                console.error('Error fetching placeholders:', error)
            })
    }

    function populateDynamicFields(placeholders) {
        if (placeholders.length === 0) {
            dynamicFields.innerHTML =
                '<p>No placeholders found in the template.</p>'
            return
        }
        placeholders.forEach((placeholder) => {
            const label = document.createElement('label')
            label.textContent = `${placeholder}: `
            const input = document.createElement('input')
            input.name = placeholder
            input.placeholder = `Enter ${placeholder}`
            dynamicFields.appendChild(label)
            dynamicFields.appendChild(input)
            dynamicFields.appendChild(document.createElement('br'))
        })
    }

    // Handle template selection change
    templateSelect.addEventListener('change', () => {
        fetchPlaceholders(templateSelect.value)
        generatePreview() // Generate preview whenever template is changed
    })

    // Handle preview button (Allow preview without filling data)
    previewBtn.addEventListener('click', () => {
        formError.textContent = '' // Clear any previous error
        generatePreview() // Allow preview even if the form is not filled
    })

    // Generate preview
    function generatePreview() {
        const formData = Array.from(
            dynamicFields.querySelectorAll('input')
        ).reduce((data, input) => {
            data[input.name] = input.value
            return data
        }, {})

        fetch('/preview-template', {
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
                    licenseKey: 'non-commercial-and-evaluation',
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
            formError.textContent = messages[langSelect.value].formError // Display the form error message
            alert(messages[langSelect.value].formError) // Show an alert asking the user to fill all fields
        } else {
            fetch('/fill-template', {
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
        }
    })

    // Change language when user selects different option
    langSelect.addEventListener('change', () => {
        updateUI(langSelect.value)
    })

    // Initial setup
    updateUI('gu')
    fetchTemplates()
})
