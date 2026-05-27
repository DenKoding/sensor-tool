let sensors = [];
let editIndex = null;

// -------------------- Speicher --------------------

function saveSensors() {
    localStorage.setItem("sensors", JSON.stringify(sensors));
}

function loadSensors() {
    const saved = localStorage.getItem("sensors");
    if (!saved) return;

    sensors = JSON.parse(saved).map(s => ({
        ...s,
        startDate: new Date(s.startDate),
        endDate: s.endDate ? new Date(s.endDate) : null
    }));
}

// -------------------- Berechnungen --------------------

function updateEndDates() {
    sensors.sort((a, b) => a.startDate - b.startDate);

    sensors.forEach((s, i) => {
        if (i < sensors.length - 1) {
            const next = sensors[i + 1];
            s.endDate = next.startDate;
            s.days = Math.round((s.endDate - s.startDate) / 86400000);
        } else {
            s.endDate = null;
            s.days = null;
        }
    });
}

function updateYearFilter() {
    const select = document.getElementById("yearFilter");
    const years = [...new Set(sensors.map(s => s.startDate.getFullYear()))].sort((a, b) => b - a);

    const current = select.value || "all";
    select.innerHTML = '<option value="all">Alle</option>';

    years.forEach(y => {
        const opt = document.createElement("option");
        opt.value = y;
        opt.textContent = y;
        select.appendChild(opt);
    });

    if ([...select.options].some(o => o.value == current)) {
        select.value = current;
    }
}

// -------------------- UI --------------------

function toggleDetails(idx) {
    const details = document.getElementById(`details-${idx}`);
    const arrow = document.getElementById(`arrow-${idx}`);

    const open = details.style.display === "block";
    details.style.display = open ? "none" : "block";
    arrow.textContent = open ? "▶" : "▼";
}

function openNewSensorForm() {
    document.getElementById("sensorForm").reset();
    document.getElementById("statusFields").style.display = "none";
    editIndex = null;
}

// -------------------- CRUD --------------------

function addSensor() {
    const startVal = document.getElementById("startDate").value;
    if (!startVal) return;

    const sensor = {
        startDate: new Date(startVal),
        lot: document.getElementById("lot").value,
        serial: document.getElementById("serial").value,
        prodNr: document.getElementById("prodNr").value,
        coupling: document.getElementById("coupling").value,
        notes: document.getElementById("notes").value,
        reklamiert: document.getElementById("reklamiert").value,
        erstattet: document.getElementById("erstattet").value,
        retour: document.getElementById("retour").value,
        endDate: null,
        days: null
    };

    if (editIndex !== null) {
        sensors[editIndex] = sensor;
        editIndex = null;
    } else {
        sensors.push(sensor);
    }

    updateEndDates();
    updateYearFilter();
    renderSensors();
    saveSensors();
    openNewSensorForm();
}

function clearAll() {
    if (!confirm("Wirklich ALLE Sensoren löschen?")) return;

    sensors = [];
    localStorage.removeItem("sensors");

    renderSensors();
    updateYearFilter();
    openNewSensorForm();
}


function editSensor(id) {
    const s = sensors[id];
    editIndex = id;

    document.getElementById("startDate").value = s.startDate.toISOString().split("T")[0];
    document.getElementById("lot").value = s.lot;
    document.getElementById("serial").value = s.serial;
    document.getElementById("prodNr").value = s.prodNr;
    document.getElementById("coupling").value = s.coupling;
    document.getElementById("notes").value = s.notes;

    document.getElementById("reklamiert").value = s.reklamiert;
    document.getElementById("erstattet").value = s.erstattet;
    document.getElementById("retour").value = s.retour;

    document.getElementById("statusFields").style.display = "block";
}

// -------------------- Rendering --------------------

function renderSensors() {
    const list = document.getElementById("sensorList");
    list.innerHTML = "";

    const filterYear = document.getElementById("yearFilter").value;
    const filterWarning = document.getElementById("warningFilter").value;

    // Sortierung: neueste zuerst
    const ordered = sensors
        .map((s, idx) => ({ s, idx }))
        .sort((a, b) => b.s.startDate - a.s.startDate);

    let visibleCount = 0; // für die Summenanzeige

    ordered.forEach(({ s, idx }) => {

        // 1) Jahresfilter
        if (filterYear !== "all" && filterYear != s.startDate.getFullYear()) return;

        // 2) Warnfilter (< 8 Tage)
        const isWarning = s.days !== null && s.days < 8;
        if (filterWarning === "warning" && !isWarning) return;

        // 3) Wenn Sensor angezeigt wird → Counter erhöhen
        visibleCount++;

        // 4) HTML erzeugen
        const entry = document.createElement("div");
        entry.className = "sensor-entry" + (isWarning ? " warning-row" : "");

        entry.innerHTML = `
            <div class="sensor-header" onclick="toggleDetails(${idx})">
                <span>${s.startDate.toLocaleDateString()}</span>
                <span>${s.coupling || "-"}</span>
                <span>${s.days !== null ? s.days + " Tage" : "-"}</span>
                <span id="arrow-${idx}" class="arrow">▶</span>
            </div>

            <div id="details-${idx}" class="details">
                <p><strong>LOT:</strong> ${s.lot || "-"}</p>
                <p><strong>Seriennummer:</strong> ${s.serial || "-"}</p>
                <p><strong>Produktionsnummer:</strong> ${s.prodNr || "-"}</p>
                <p><strong>Bemerkung:</strong> ${s.notes || "-"}</p>
                <p><strong>Reklamiert:</strong> ${s.reklamiert}</p>
                <p><strong>Erstattet:</strong> ${s.erstattet}</p>
                <p><strong>Retoure:</strong> ${s.retour}</p>

                <button onclick="event.stopPropagation(); editSensor(${idx})">Bearbeiten</button>
            </div>
        `;

        list.appendChild(entry);
    });

    // 5) Summenanzeige aktualisieren
    document.getElementById("sensorSummary").textContent =
        "Erfasste Sensoren: " + visibleCount;
}


// -------------------- Excel Import --------------------

function triggerExcelImport() {
    document.getElementById("excelInput").click();
}

function excelImport(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array", cellStyles: true });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const range = XLSX.utils.decode_range(sheet['!ref']);

        for (let row = range.s.r + 1; row <= range.e.r; row++) {

            const getCell = (col) => sheet[XLSX.utils.encode_cell({ r: row, c: col })];

            const cA = getCell(0); // Einsatzdatum
            const cB = getCell(1); // LOT
            const cC = getCell(2); // SN
            const cD = getCell(3); // (01)
            const cE = getCell(4); // Kopplungscode
            const cG = getCell(6); // Bemerkung
            const cH = getCell(7); // Reklamiert (Farbe)
            const cI = getCell(8); // Erstattet (Farbe)
            const cJ = getCell(9); // Retoure (X)

            // Leere Zeilen überspringen
            if (!cA || !cA.v) continue;

            // Excel-Datum konvertieren
            const startDate = XLSX.SSF.parse_date_code(cA.v);
            const jsDate = new Date(startDate.y, startDate.m - 1, startDate.d);

            // Farben auslesen
            const getColor = (cell) => {
                if (!cell || !cell.s || !cell.s.fgColor) return "nein";
                const rgb = cell.s.fgColor.rgb;
                if (!rgb) return "nein";

                if (rgb.toUpperCase() === "FF8ED973") return "ja";   // grün
                if (rgb.toUpperCase() === "FFFFC7CE") return "nein"; // rot

                return "nein";
            };

            const reklamiert = cH && cH.v && String(cH.v).trim().toLowerCase() === "x" ? "ja" : "nein";
			const erstattet  = cI && cI.v && String(cI.v).trim().toLowerCase() === "x" ? "ja" : "nein";
			const retour     = cJ && cJ.v && String(cJ.v).trim().toLowerCase() === "x" ? "ja" : "nein";

            sensors.push({
                startDate: jsDate,
                lot: cB ? cB.v : "",
                serial: cC ? cC.v : "",
                prodNr: cD ? cD.v : "",
                coupling: cE ? cE.v : "",
                notes: cG ? cG.v : "",
                reklamiert,
                erstattet,
                retour,
                endDate: null,
                days: null
            });
        }

        updateEndDates();
        updateYearFilter();
        renderSensors();
        saveSensors();
    };

    reader.readAsArrayBuffer(file);
}

// -------------------- Backup Import/Export --------------------

function backupExport() {
    const blob = new Blob([JSON.stringify(sensors, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "sensor-backup.json";
    a.click();

    URL.revokeObjectURL(url);
}

function triggerBackupImport() {
    document.getElementById("backupInput").click();
}

function backupImport(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        sensors = JSON.parse(e.target.result).map(s => ({
            ...s,
            startDate: new Date(s.startDate),
            endDate: s.endDate ? new Date(s.endDate) : null
        }));

        updateEndDates();
        updateYearFilter();
        renderSensors();
        saveSensors();
    };

    reader.readAsText(file);
}

let scannerStream = null;
let codeReader = null;

async function openScanner() {
    document.getElementById("scannerOverlay").style.display = "flex";

    const { BrowserMultiFormatReader } = ZXing;

    codeReader = new BrowserMultiFormatReader();

    try {
        scannerStream = await navigator.mediaDevices.getUserMedia({
            video: { facingMode: "environment" }
        });

        const video = document.getElementById("scannerVideo");
        video.srcObject = scannerStream;

        codeReader.decodeFromVideoDevice(null, "scannerVideo", (result, err) => {
            if (result) {
                handleScannedCode(result.getText());
                closeScanner();
            }
        });

    } catch (err) {
        alert("Kamera konnte nicht geöffnet werden.");
        closeScanner();
    }
}

function closeScanner() {
    document.getElementById("scannerOverlay").style.display = "none";

    if (codeReader) {
        codeReader.reset();
        codeReader = null;
    }

    if (scannerStream) {
        scannerStream.getTracks().forEach(t => t.stop());
        scannerStream = null;
    }
}

function handleScannedCode(text) {
    console.log("GESCAannter Code:", text);
    alert("Gescannter Code:\n\n" + text);

    const extract = (ai) => {
        const regex = new RegExp(`\\(${ai}\\)([^()]+)`);
        const match = text.match(regex);
        return match ? match[1] : "";
    };

    const prodNr = extract("01");
    const lot = extract("10");
    const serial = extract("21");
    const coupling = extract("91");

    document.getElementById("prodNr").value = prodNr;
    document.getElementById("lot").value = lot;
    document.getElementById("serial").value = serial;
    document.getElementById("coupling").value = coupling;

    alert("Sensor-Daten erfolgreich gescannt!");
}

function triggerImageUpload() {
    document.getElementById("imageInput").click();
}

function closeImageOverlay() {
    document.getElementById("imageOverlay").style.display = "none";
}

async function decodeImage(event) {
    const file = event.target.files[0];
    if (!file) return;

    const img = document.getElementById("previewImage");
    img.src = URL.createObjectURL(file);

    document.getElementById("imageOverlay").style.display = "flex";

    const { BrowserMultiFormatReader } = ZXing;
    const reader = new BrowserMultiFormatReader();

    try {
        const result = await reader.decodeFromImage(img);
        handleScannedCode(result.getText());
        closeImageOverlay();
    } catch (err) {
        alert("Kein Data-Matrix-Code erkannt. Bitte ein schärferes Foto versuchen.");
    }
}

// -------------------- Init --------------------

window.onload = function () {
    loadSensors();
    updateEndDates();
    updateYearFilter();
    openNewSensorForm();
    renderSensors();
};
