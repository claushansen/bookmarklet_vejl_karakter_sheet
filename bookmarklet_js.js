
require.config({
    paths: {
        xlsx: "https://cdn.jsdelivr.net/gh/sheetjs/sheetjs/dist/xlsx.full.min"
    }
});

require(["xlsx"], function (XLSX) {
    console.log('SheetJS (XLSX) loaded..');

    var elever = $('#participants tbody tr:not(.emptyrow)')
        .has('input.usercheckbox:checked')
        .find('th.c1');
    if (elever.length === 0) {
        alert('Ingen elever fundet eller ingen elever valgt!');
        return;
    }

    // Overskriftsrækken med ekstra kolonne til Gennemsnit
    var header = ["Navn", "Dansk Mundligt", "Dansk Skriftligt", "Engelsk Mundtligt", "Engelsk Skriftligt", "Informatik", "Samfundsfag",
        "Markedskommunikation", "Virksomhedsøkonomi", "Erhvervsjura",
        "Gennemsnit"];
    var sheet1Data = [header];
    var sheet2Data = [header];

    elever.each(function (index) {
        var linkElement = this.querySelector('a');
        var fuldtNavn = "";
        if (linkElement) {
            fuldtNavn = Array.from(linkElement.childNodes)
                .filter(node => node.nodeType === Node.TEXT_NODE)
                .map(node => node.textContent.trim())
                .join("");
        } else {
            fuldtNavn = $(this).text().trim();
        }
        // Excel-rækkenummer (den første elevrække er 2, fordi overskriften er i række 1)
        var excelRowNum = index + 2;

        // Benyt IFERROR: AVERAGE(...) → IFERROR(AVERAGE(...), "")
        var row = [
            fuldtNavn,
            "", "", "", "", "", "", "", "", "",
            { f: "IFERROR(AVERAGE(B" + excelRowNum + ":J" + excelRowNum + "), \"\")" }
        ];
        sheet1Data.push(row);
        sheet2Data.push(row);
    });

    var wb = XLSX.utils.book_new();
    var ws1 = XLSX.utils.aoa_to_sheet(sheet1Data);
    var ws2 = XLSX.utils.aoa_to_sheet(sheet2Data);
    XLSX.utils.book_append_sheet(wb, ws1, "Halvår 1");
    XLSX.utils.book_append_sheet(wb, ws2, "Halvår 2");

    var wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }

    var year = new Date().getFullYear();
    var filename = "Vejledende Karakterer - " + year + ".xlsx";

    if (typeof saveAs === "undefined") {
        var fsScript = document.createElement('script');
        fsScript.src = "https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/1.3.8/FileSaver.min.js";
        fsScript.onload = function () {
            saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), filename);
        };
        document.head.appendChild(fsScript);
    } else {
        saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), filename);
    }
});

