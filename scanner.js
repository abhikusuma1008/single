let scanner = null;

document.getElementById('scan-btn').addEventListener('click', function() {
    if (!scanner) {
        scanner = new Instascan.Scanner({ video: document.getElementById('qr-video') });

        scanner.addListener('scan', function (content) {
            document.getElementById('qr-result').innerText = 'Scanned: ' + content;
            saveToExcel(content);
        });

        Instascan.Camera.getCameras().then(function (cameras) {
            if (cameras.length > 0) {
                scanner.start(cameras[0]);
            } else {
                console.error('No cameras found.');
            }
        }).catch(function (e) {
            console.error(e);
        });
    } else {
        scanner.start();
    }
});

function saveToExcel(content) {
    // Create Excel workbook
    let wb = XLSX.utils.book_new();

    // Create worksheet
    let ws = XLSX.utils.json_to_sheet([{ QR_Code: content }]);

    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Scanned QR Codes');

    // Generate a unique filename
    let fileName = 'scanned_qr_codes_' + Date.now() + '.xlsx';

    // Save workbook
    XLSX.writeFile(wb, fileName);
}
