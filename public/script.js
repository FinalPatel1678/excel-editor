document.addEventListener('DOMContentLoaded', () => {
    const templateSelect = document.getElementById('template')
    const dynamicFields = document.getElementById('dynamicFields')
    const previewBtn = document.getElementById('previewBtn')
    const downloadBtn = document.getElementById('downloadBtn')
    const previewTable = document.getElementById('previewTable')

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
        console.log('Selected template:', selectedFile)

        if (!selectedFile) {
            console.log('No template selected.')
            return
        }

        // Call the backend to fetch placeholders for the selected template
        fetch(`/get-placeholders?file=${selectedFile}`)
            .then((res) => res.json())
            .then((data) => {
                console.log('Placeholders fetched:', data.placeholders)

                // Generate form fields dynamically
                dynamicFields.innerHTML = '' // Clear existing fields
                data.placeholders.forEach((placeholder) => {
                    const label = document.createElement('label')
                    label.textContent = `${placeholder}: `
                    const input = document.createElement('input')
                    input.name = placeholder // Field name will be the placeholder
                    input.placeholder = `Enter ${placeholder}`
                    dynamicFields.appendChild(label)
                    dynamicFields.appendChild(input)
                    dynamicFields.appendChild(document.createElement('br'))
                })
            })
            .catch((error) => {
                console.error('Error fetching placeholders:', error)
                alert('Failed to load placeholders. Please try again.')
            })
    })

    // Handle preview button
    previewBtn.addEventListener('click', () => {
        const formData = new FormData()
        dynamicFields.querySelectorAll('input').forEach((input) => {
            formData.append(input.name, input.value)
        })

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
            .then((res) => res.blob())
            .then((blob) => {
                const url = window.URL.createObjectURL(blob)
                const iframe = document.createElement('iframe')
                iframe.src = url
                iframe.width = '100%'
                iframe.height = '500px'
                previewTable.innerHTML = '' // Clear previous preview
                previewTable.appendChild(iframe)
            })
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
                a.download = `filled_${data.fileName}`
                document.body.appendChild(a)
                a.click()
                a.remove()
            })
    })
})
