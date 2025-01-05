document.addEventListener('DOMContentLoaded', () => {
    const templateSelect = document.getElementById('template')
    const dynamicFields = document.getElementById('dynamicFields')
    const previewBtn = document.getElementById('previewBtn')
    const downloadBtn = document.getElementById('downloadBtn')
    const excelPreview = document.getElementById('excelPreview')

    let handsontableInstance // Handsontable instance

    // Fetch available templates when the page loads
    fetch('/get-templates')
        .then((res) => res.json())
        .then((data) => {
            data.templates.forEach((template) => {
                const option = document.createElement('option')
                option.value = template
                option.textContent = template
                templateSelect.appendChild(option)
            })

            // Trigger change event to fetch placeholders for the selected template
            if (templateSelect.value) {
                templateSelect.dispatchEvent(new Event('change'))
            }
        })
        .catch((error) => console.error('Error fetching templates:', error))

    // Fetch placeholders when a template is selected
    templateSelect.addEventListener('change', () => {
        const selectedFile = templateSelect.value

        if (!selectedFile) {
            return
        }

        // Call the backend to fetch placeholders for the selected template
        fetch(`/get-placeholders?file=${selectedFile}`)
            .then((res) => res.json())
            .then((data) => {
                dynamicFields.innerHTML = '' // Clear existing fields
                data.placeholders.forEach((placeholder) => {
                    const label = document.createElement('label')
                    label.textContent = `${placeholder}: `
                    const input = document.createElement('input')
                    input.name = placeholder
                    input.placeholder = `Enter ${placeholder}`
                    dynamicFields.appendChild(label)
                    dynamicFields.appendChild(input)
                    dynamicFields.appendChild(document.createElement('br'))
                })
            })
    })

    // Handle preview button
    previewBtn.addEventListener('click', () => {
        const data = {
            fileName: templateSelect.value,
            formData: Object.fromEntries(
                Array.from(dynamicFields.querySelectorAll('input')).map(
                    (input) => [input.name, input.value]
                )
            ),
        }

        fetch('/preview-template', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data),
        })
            .then((res) => res.json())
            .then((excelData) => {
                // Render Excel data using Handsontable
                if (handsontableInstance) {
                    handsontableInstance.destroy() // Destroy existing instance if any
                }
                handsontableInstance = new Handsontable(excelPreview, {
                    data: excelData,
                    colHeaders: true,
                    rowHeaders: true,
                    licenseKey: 'non-commercial-and-evaluation', // For free version
                })
            })
            .catch((error) =>
                console.error('Error previewing template:', error)
            )
    })

    // Handle download button
    downloadBtn.addEventListener('click', () => {
        const data = {
            fileName: templateSelect.value,
            formData: Object.fromEntries(
                Array.from(dynamicFields.querySelectorAll('input')).map(
                    (input) => [input.name, input.value]
                )
            ),
        }

        fetch('/fill-template', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data),
        })
            .then((res) => res.blob())
            .then((blob) => {
                const url = window.URL.createObjectURL(blob)
                const a = document.createElement('a')
                a.href = url
                a.download = `${data.fileName}` // Ensure filename ends with .xlsx
                document.body.appendChild(a)
                a.click()
                a.remove()
            })
    })
})
