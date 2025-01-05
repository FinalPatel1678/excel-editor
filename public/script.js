document.addEventListener('DOMContentLoaded', () => {
    const templateSelect = document.getElementById('template')
    const dynamicFields = document.getElementById('dynamicFields')
    const previewBtn = document.getElementById('previewBtn')
    const downloadBtn = document.getElementById('downloadBtn')
    const excelPreview = document.getElementById('excelPreview')
    const formError = document.getElementById('formError')
    const templateLoading = document.getElementById('templateLoading')

    let handsontableInstance // Handsontable instance

    // Fetch available templates
    function fetchTemplates() {
        templateLoading.textContent = 'Loading templates...'
        fetch('/get-templates')
            .then((res) => res.json())
            .then((data) => {
                templateSelect.innerHTML = '' // Clear existing options
                data.templates.forEach((template) => {
                    const option = document.createElement('option')
                    option.value = template
                    option.textContent = template
                    templateSelect.appendChild(option)
                })

                templateLoading.textContent = '' // Clear loading text
                if (templateSelect.value) {
                    fetchPlaceholders(templateSelect.value)
                }
            })
            .catch((error) => {
                console.error('Error fetching templates:', error)
                templateLoading.textContent = 'Failed to load templates!'
            })
    }

    // Fetch placeholders for the selected template
    function fetchPlaceholders(template) {
        dynamicFields.innerHTML = '' // Clear existing fields
        excelPreview.innerHTML = '' // Clear existing preview

        fetch(`/get-placeholders?file=${template}`)
            .then((res) => res.json())
            .then((data) => {
                if (data.placeholders.length === 0) {
                    dynamicFields.innerHTML =
                        '<p>No placeholders found in the template.</p>'
                    return
                }

                data.placeholders.forEach((placeholder) => {
                    const label = document.createElement('label')
                    label.textContent = `${placeholder}: `
                    const input = document.createElement('input')
                    input.name = placeholder
                    input.placeholder = `Enter ${placeholder}`
                    input.required = true
                    dynamicFields.appendChild(label)
                    dynamicFields.appendChild(input)
                    dynamicFields.appendChild(document.createElement('br'))
                })
            })
            .catch((error) =>
                console.error('Error fetching placeholders:', error)
            )
    }

    // Handle template selection change
    templateSelect.addEventListener('change', () => {
        fetchPlaceholders(templateSelect.value)
    })

    // Handle preview button
    previewBtn.addEventListener('click', () => {
        formError.textContent = '' // Clear any previous error
        const formData = Object.fromEntries(
            Array.from(dynamicFields.querySelectorAll('input')).map((input) => [
                input.name,
                input.value,
            ])
        )

        if (Object.values(formData).some((value) => !value.trim())) {
            formError.textContent = 'Please fill all fields before previewing.'
            return
        }

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
            .catch((error) => console.error('Error generating preview:', error))
    })

    // Handle download button
    downloadBtn.addEventListener('click', () => {
        const formData = Object.fromEntries(
            Array.from(dynamicFields.querySelectorAll('input')).map((input) => [
                input.name,
                input.value,
            ])
        )

        fetch('/fill-template', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ fileName: templateSelect.value, formData }),
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
            .catch((error) => console.error('Error downloading file:', error))
    })

    // Initial fetch of templates
    fetchTemplates()
})
