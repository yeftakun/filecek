document.getElementById("processButton").addEventListener("click", processFile);

async function processFile() {
    const fileInput = document.getElementById('fileInput');
    const resultDiv = document.getElementById('result');
    resultDiv.innerHTML = '';

    if (!fileInput.files.length) {
        resultDiv.innerHTML = '<p class="text-danger">Pilih file terlebih dahulu!</p>';
        return;
    }

    const file = fileInput.files[0];
    const fileType = file.type;
    let characterCount = 0;
    let wordCount = 0;
    let imageCount = 0;
    let pageCount = 0;

    try {
        if (fileType === 'application/pdf') {
            const buffer = await file.arrayBuffer();
            const pdfData = new Uint8Array(buffer);

            // Load PDF using pdfjs-dist
            const pdf = await pdfjsLib.getDocument(pdfData).promise;
            pageCount = pdf.numPages; // Count number of pages
            for (let i = 0; i < pdf.numPages; i++) {
                const page = await pdf.getPage(i + 1);
                const textContent = await page.getTextContent();
                const textItems = textContent.items.map(item => item.str).join('');
                
                // Hitung jumlah karakter
                characterCount += textItems.length;

                // Hitung jumlah kata
                wordCount += textItems.split(/\s+/).filter(word => word.length > 0).length;

                // Menghitung jumlah gambar
                const ops = await page.getOperatorList();
                imageCount += ops.fnArray.filter(fn => fn === pdfjsLib.OPS.paintJpegXObject || fn === pdfjsLib.OPS.paintImageXObject).length;
            }
        } else if (fileType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer });
            const textContent = result.value;

            // Hitung jumlah karakter
            characterCount = textContent.length;

            // Hitung jumlah kata
            wordCount = textContent.split(/\s+/).filter(word => word.length > 0).length;

            // Menghitung jumlah halaman (approximate based on sections)
            const doc = await mammoth.convertToHtml({ arrayBuffer });
            pageCount = doc.value.match(/<section/g)?.length || 1;

            // Menghitung jumlah gambar
            const parser = new DOMParser();
            const docImages = parser.parseFromString(doc.value, 'text/html').querySelectorAll('img');
            imageCount = docImages.length;
        } else if (fileType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            let textContent = '';

            // Count each sheet as a "page"
            pageCount = workbook.SheetNames.length;
            
            // Baca setiap sheet dalam workbook
            workbook.SheetNames.forEach(sheetName => {
                const sheet = workbook.Sheets[sheetName];
                const sheetText = XLSX.utils.sheet_to_csv(sheet);
                textContent += sheetText;
            });

            // Hitung jumlah karakter dan kata di Excel
            characterCount = textContent.length;
            wordCount = textContent.split(/\s+/).filter(word => word.length > 0).length;

            // Saat ini, Excel tidak memiliki gambar yang dapat dibaca secara langsung dari SheetJS.
            imageCount = 0;
        } else if (fileType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation') {
            const arrayBuffer = await file.arrayBuffer();
            const zip = new JSZip();
            await zip.loadAsync(arrayBuffer);
            const slideRegex = /ppt\/slides\/slide\d+\.xml/;
            let textContent = '';

            // Count each slide as a "page"
            pageCount = Object.keys(zip.files).filter(path => slideRegex.test(path)).length;

            // Iterasi setiap file slide di dalam PPTX
            for (let relativePath in zip.files) {
                if (slideRegex.test(relativePath)) {
                    const slideText = await zip.files[relativePath].async("text");
                    const matches = slideText.match(/<a:t>(.*?)<\/a:t>/g);

                    // Ekstrak teks dari tag <a:t>
                    if (matches) {
                        matches.forEach(match => {
                            const text = match.replace(/<\/?a:t>/g, '');
                            textContent += text;
                        });
                    }

                    // Menghitung gambar berdasarkan tag XML gambar dalam slide
                    imageCount += (slideText.match(/<p:blipFill/g) || []).length;
                }
            }

            // Hitung jumlah karakter dan kata di PowerPoint
            characterCount = textContent.length;
            wordCount = textContent.split(/\s+/).filter(word => word.length > 0).length;
        } else {
            resultDiv.innerHTML = '<p class="text-danger">Format file tidak didukung!</p>';
            return;
        }

        resultDiv.innerHTML = `
            <h3>Hasil</h3>
            <p>Jumlah Halaman: <strong>${pageCount}</strong></p>
            <p>Jumlah Karakter: <strong>${characterCount}</strong></p>
            <p>Jumlah Kata: <strong>${wordCount}</strong></p>
            <p>Jumlah Gambar: <strong>${imageCount}</strong></p>
        `;
    } catch (error) {
        console.error("Error processing file:", error);
        resultDiv.innerHTML = '<p class="text-danger">Terjadi kesalahan saat memproses file.</p>';
    }
}


document.addEventListener("DOMContentLoaded", function () {
    const dropZone = document.getElementById("dropZone");
    const fileInput = document.getElementById("fileInput");
    const processButton = document.getElementById("processButton");

    // Trigger file input when drop zone is clicked
    dropZone.addEventListener("click", () => fileInput.click());

    // Handle file selection
    fileInput.addEventListener("change", handleFiles);
    
    // Drag-and-drop events
    dropZone.addEventListener("dragover", (event) => {
        event.preventDefault();
        dropZone.classList.add("dragging");
    });

    dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragging"));

    dropZone.addEventListener("drop", (event) => {
        event.preventDefault();
        dropZone.classList.remove("dragging");
        const files = event.dataTransfer.files;
        if (files.length) {
            fileInput.files = files; // Set the file input to the dropped files
            handleFiles(); // Call the handler function
        }
    });

    // Function to process selected file(s)
    function handleFiles() {
        const files = fileInput.files;
        if (files.length > 0) {
            const file = files[0];
            document.getElementById("result").textContent = `File "${file.name}" siap untuk diproses.`;
        }
    }

    // Process button action
    processButton.addEventListener("click", () => {
        if (fileInput.files.length === 0) {
            alert("Silakan unggah file terlebih dahulu!");
            return;
        }
        // Implement file processing logic here
        document.getElementById("result").textContent = "Sedang memproses file...";
    });
});
