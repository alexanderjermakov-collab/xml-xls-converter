const TARGET_COLUMNS = [5, 8, 11, 14];
const XLS_HEADER_ROW_CANDIDATES = [24, 25];
const MAP_HEADER_SCAN_LIMIT = 100;
const TARGET_XLS_SHEET_NAME = "AQ_Tbl";
const TARGET_MAP_SHEET_NAME = "MAP";
const REQUIRED_MAP_HEADERS = ["DSP", "Description", "Parameters", "Mapped to DAP XML"];
const OUTPUT_HEADERS = ["DSP", "Description", "Parameters", "Column", "Previous Value", "New Value", "XML Mapping"];

const state = {
  xmlFile: null,
  xlsFile: null,
  mapFile: null,
  workbook: null,
  outputName: "",
  logText: "",
  logName: ""
};

const elements = {
  versionInput: document.getElementById("versionInput"),
  xmlInput: document.getElementById("xmlInput"),
  xlsInput: document.getElementById("xlsInput"),
  mapInput: document.getElementById("mapInput"),
  outputNameInput: document.getElementById("outputNameInput"),
  convertButton: document.getElementById("convertButton"),
  downloadXlsButton: document.getElementById("downloadXlsButton"),
  logOutput: document.getElementById("logOutput"),
  logMeta: document.getElementById("logMeta"),
  statusText: document.getElementById("statusText"),
  progressFill: document.getElementById("progressFill"),
  progressValue: document.getElementById("progressValue"),
  progressTrack: document.querySelector(".progress-track"),
  errorText: document.getElementById("errorText"),
  summaryText: document.getElementById("summaryText"),
  bottomMonitor: document.querySelector(".bottom-monitor")
};

function normalise(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[\s_\-./\\()[\]{}:;]+/g, "");
}

function cellAddress(rowIndex, columnNumber) {
  return XLSX.utils.encode_cell({ r: rowIndex, c: columnNumber - 1 });
}

function isRelevant(value) {
  const text = String(value ?? "").trim();
  return text !== "" && normalise(text) !== "notrelevant";
}

function setStatus(message, summary = "", isError = false, debugDetails = "") {
  elements.statusText.textContent = message;
  elements.summaryText.textContent = summary;
  elements.bottomMonitor.classList.toggle("error", isError);
  elements.errorText.textContent = debugDetails || (isError ? message : "No errors.");
}

function setProgress(value) {
  const progress = Math.max(0, Math.min(100, Number(value) || 0));
  elements.progressFill.style.width = `${progress}%`;
  elements.progressValue.textContent = `${progress}%`;
  elements.progressTrack.setAttribute("aria-valuenow", String(progress));
}

function setFile(targetId, file) {
  if (!file) return;

  if (targetId === "xmlInput") state.xmlFile = file;
  if (targetId === "xlsInput") state.xlsFile = file;
  if (targetId === "mapInput") state.mapFile = file;
  if (targetId === "outputNameInput") {
    elements.outputNameInput.value = file.name.replace(/\.(xls|xlsx|xlsm|csv)$/i, "");
  }

  const zone = document.querySelector(`.drop-zone[data-target="${targetId}"]`);
  const label = zone?.querySelector("span");
  if (label && targetId !== "outputNameInput") label.textContent = file.name;
  if (label && targetId === "outputNameInput") label.textContent = elements.outputNameInput.value;
  updateOutputNameSuggestion();
  setStatus("Files updated.", readinessSummary());
}

function readinessSummary() {
  const ready = [
    state.xmlFile ? "XML" : null,
    state.xlsFile ? "XLS" : null,
    state.mapFile ? "MAP" : null,
    elements.versionInput.value.trim() ? "version" : null
  ].filter(Boolean);
  return ready.length ? `${ready.join(", ")} ready` : "";
}

function updateOutputNameSuggestion() {
  const outputZone = document.querySelector('.drop-zone[data-target="outputNameInput"] span');
  const typedName = elements.outputNameInput.value.trim();
  if (typedName) {
    if (outputZone) outputZone.textContent = typedName;
    return;
  }

  if (!state.xlsFile) return;
  const version = elements.versionInput.value.trim().replace(/[\\/:*?"<>|]+/g, "_");
  const suggested = version ? `${baseName(state.xlsFile.name)}_${version}_converted` : `${baseName(state.xlsFile.name)}_converted`;
  if (outputZone) outputZone.textContent = `${suggested}.xlsx`;
}

function wireDropZones() {
  document.querySelectorAll(".drop-zone").forEach((zone) => {
    const targetId = zone.dataset.target;
    const input = document.getElementById(targetId);
    const button = zone.querySelector("button");

    if (button && input) {
      button.addEventListener("click", (event) => {
        event.preventDefault();
        input.click();
      });
    }

    if (input?.type === "file") {
      input.addEventListener("change", () => setFile(targetId, input.files[0]));
    }

    zone.addEventListener("dragover", (event) => {
      event.preventDefault();
      zone.classList.add("is-over");
    });

    zone.addEventListener("dragleave", () => zone.classList.remove("is-over"));

    zone.addEventListener("drop", (event) => {
      event.preventDefault();
      zone.classList.remove("is-over");
      setFile(targetId, event.dataTransfer.files[0]);
    });
  });
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsText(file);
  });
}

async function readWorkbook(file) {
  const buffer = await readFileAsArrayBuffer(file);
  return XLSX.read(buffer, { type: "array", cellDates: true, cellNF: true });
}

function sheetRange(sheet) {
  if (!sheet["!ref"]) throw new Error("Workbook sheet is empty.");
  return XLSX.utils.decode_range(sheet["!ref"]);
}

function sheetNames(workbook) {
  return workbook.SheetNames || [];
}

function workbookSheetInfo(workbook) {
  const metadata = workbook.Workbook?.Sheets || [];
  return sheetNames(workbook).map((name, index) => ({
    name,
    hidden: Boolean(metadata[index]?.Hidden)
  }));
}

function findSheetName(workbook, preferredName) {
  const names = sheetNames(workbook);
  const exact = names.find((name) => name === preferredName);
  if (exact) return exact;

  const normalisedPreferred = normalise(preferredName);
  return names.find((name) => normalise(name) === normalisedPreferred) || "";
}

function getRequiredSheet(workbook, preferredName, label) {
  const matchedName = findSheetName(workbook, preferredName);
  if (!matchedName) {
    throw new Error(`${label} worksheet "${preferredName}" was not found. Available worksheets: ${sheetNames(workbook).join(", ") || "none"}.`);
  }

  return { sheet: workbook.Sheets[matchedName], sheetName: matchedName };
}

function candidateSheetNames(workbook) {
  const info = workbookSheetInfo(workbook);
  const visibleNames = info.filter((item) => !item.hidden).map((item) => item.name);
  return visibleNames.length ? visibleNames : sheetNames(workbook);
}

function readRow(sheet, rowIndex, lastColumnIndex) {
  const row = [];
  for (let columnIndex = 0; columnIndex <= lastColumnIndex; columnIndex += 1) {
    const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
    row.push(sheet[address]?.v ?? "");
  }
  return row;
}

function headerIndexes(row) {
  const indexes = {};
  row.forEach((header, columnIndex) => {
    const key = normalise(header);
    if (key && indexes[key] === undefined) indexes[key] = columnIndex;
  });
  return indexes;
}

function rowContainsHeaders(row, requiredHeaders) {
  const normalisedRow = row.map(normalise);
  return requiredHeaders.every((header) => normalisedRow.includes(normalise(header)));
}

function findHeaderRow(sheet, requiredHeaders, options = {}) {
  const range = sheetRange(sheet);
  const lastColumnIndex = Math.max(range.e.c, 20);
  const label = options.label || "Workbook";

  if (options.fixedHeaderRowNumbers?.length) {
    for (const fixedHeaderRowNumber of options.fixedHeaderRowNumbers) {
      const rowIndex = fixedHeaderRowNumber - 1;
      const row = readRow(sheet, rowIndex, lastColumnIndex);

      if (rowContainsHeaders(row, requiredHeaders)) {
        return { rowIndex, rowNumber: fixedHeaderRowNumber, indexes: headerIndexes(row), range };
      }
    }

    throw new Error(`${label} header row must be ${options.fixedHeaderRowNumbers.join(" or ")} and contain required columns: ${requiredHeaders.join(", ")}.`);
  }

  if (options.fixedHeaderRowNumber) {
    const rowIndex = options.fixedHeaderRowNumber - 1;
    const row = readRow(sheet, rowIndex, lastColumnIndex);

    if (!rowContainsHeaders(row, requiredHeaders)) {
      throw new Error(`${label} header row ${options.fixedHeaderRowNumber} must contain required columns: ${requiredHeaders.join(", ")}.`);
    }

    return { rowIndex, rowNumber: options.fixedHeaderRowNumber, indexes: headerIndexes(row), range };
  }

  const scanEndRow = Math.min(range.e.r, MAP_HEADER_SCAN_LIMIT - 1);
  for (let rowIndex = range.s.r; rowIndex <= scanEndRow; rowIndex += 1) {
    const row = readRow(sheet, rowIndex, lastColumnIndex);
    if (rowContainsHeaders(row, requiredHeaders)) {
      return { rowIndex, rowNumber: rowIndex + 1, indexes: headerIndexes(row), range };
    }
  }

  throw new Error(`${label} header row was not found in the first ${MAP_HEADER_SCAN_LIMIT} rows. Required columns: ${requiredHeaders.join(", ")}.`);
}

function findHeaderRowInWorkbook(workbook, headers, options = {}) {
  const preferredSheetName = options.sheetName;
  const label = options.label || "Workbook";

  if (preferredSheetName) {
    const matchedName = findSheetName(workbook, preferredSheetName);
    if (matchedName) {
      const selected = { sheet: workbook.Sheets[matchedName], sheetName: matchedName };
      try {
        const header = findHeaderRow(selected.sheet, headers, options);
        return { ...selected, header, warning: matchedName === preferredSheetName ? "" : `Used worksheet "${matchedName}" as a name match for required worksheet "${preferredSheetName}".` };
      } catch (error) {
        if (!options.allowSheetFallback) {
          throw new Error(`${label} worksheet "${selected.sheetName}": ${error.message}`);
        }
      }
    } else if (!options.allowSheetFallback) {
      getRequiredSheet(workbook, preferredSheetName, label);
    }

    if (options.allowSheetFallback) {
      const tried = matchedName ? [matchedName] : [];
      for (const sheetName of candidateSheetNames(workbook)) {
        if (tried.includes(sheetName)) continue;
        try {
          const sheet = workbook.Sheets[sheetName];
          const header = findHeaderRow(sheet, headers, options);
          return {
            sheet,
            sheetName,
            header,
            warning: `Required worksheet "${preferredSheetName}" was not usable. Used worksheet "${sheetName}" because it contains the required headers.`
          };
        } catch (_) {
          // Keep scanning candidate sheets.
        }
      }

      throw new Error(`${label} worksheet "${preferredSheetName}" was not found with required headers. Available worksheets: ${sheetNames(workbook).join(", ") || "none"}.`);
    }
  }

  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const header = findHeaderRow(sheet, headers, options);
  return { sheet, sheetName: firstSheetName, header };
}

function getCellValue(row, indexes, header) {
  return row[indexes[normalise(header)]] ?? "";
}

function setSheetCell(sheet, rowIndex, columnNumber, value, previousCell) {
  const address = cellAddress(rowIndex, columnNumber);
  const newCell = { t: "s", v: String(value) };
  if (previousCell?.z) newCell.z = previousCell.z;
  sheet[address] = { ...(previousCell || {}), ...newCell };
}

function extractRows(workbook, headers, options = {}) {
  const selected = findHeaderRowInWorkbook(workbook, headers, options);
  const sheet = selected.sheet;
  const header = selected.header;
  const rows = [];
  const dataStartRowIndex = header.rowIndex + 1;

  for (let rowIndex = dataStartRowIndex; rowIndex <= header.range.e.r; rowIndex += 1) {
    const row = readRow(sheet, rowIndex, Math.max(header.range.e.c, 20));
    const hasData = row.some((value) => String(value ?? "").trim() !== "");
    if (!hasData) continue;

    rows.push({
      rowIndex,
      row,
      dsp: getCellValue(row, header.indexes, "DSP"),
      description: getCellValue(row, header.indexes, "Description"),
      parameters: getCellValue(row, header.indexes, "Parameters"),
      mappedXml: getCellValue(row, header.indexes, "Mapped to DAP XML")
    });
  }

  return { sheet, sheetName: selected.sheetName, header, rows };
}

function buildTargetRowIndex(targetRows) {
  const index = new Map();
  targetRows.forEach((item) => {
    const key = rowKey(item.dsp, item.description, item.parameters);
    if (!index.has(key)) index.set(key, item);
  });
  return index;
}

function rowKey(dsp, description, parameters) {
  return [dsp, description, parameters].map(normalise).join("|");
}

function nodePath(element) {
  const names = [];
  let current = element;
  while (current && current.nodeType === Node.ELEMENT_NODE) {
    names.unshift(current.localName || current.nodeName);
    current = current.parentElement;
  }
  return names.join("/");
}

function directText(element) {
  const textNodes = Array.from(element.childNodes).filter((node) => node.nodeType === Node.TEXT_NODE);
  return textNodes.map((node) => node.nodeValue.trim()).filter(Boolean).join(" ").trim();
}

function preferredElementValue(element, matchedAttributeName = "") {
  const valueAttributes = ["value", "val", "hex", "gain", "data", "default"];
  for (const name of valueAttributes) {
    if (element.hasAttribute(name)) return element.getAttribute(name);
  }
  if (matchedAttributeName) return element.getAttribute(matchedAttributeName);
  return directText(element) || element.textContent.trim();
}

function elementContext(element) {
  const pieces = [];
  let current = element;
  while (current && current.nodeType === Node.ELEMENT_NODE) {
    pieces.push(current.outerHTML.slice(0, 1200));
    current = current.parentElement;
  }
  return normalise(pieces.join(" "));
}

function buildXmlCandidates(xmlDocument) {
  const candidates = [];
  const elements = Array.from(xmlDocument.getElementsByTagName("*"));

  elements.forEach((element) => {
    const attributes = Array.from(element.attributes || []);
    const path = nodePath(element);
    const tagName = element.localName || element.nodeName;

    candidates.push({
      key: normalise(tagName),
      rawKey: tagName,
      value: preferredElementValue(element),
      path,
      context: elementContext(element)
    });

    candidates.push({
      key: normalise(path),
      rawKey: path,
      value: preferredElementValue(element),
      path,
      context: elementContext(element)
    });

    attributes.forEach((attribute) => {
      candidates.push({
        key: normalise(attribute.value),
        rawKey: attribute.value,
        value: preferredElementValue(element, attribute.name),
        path,
        context: elementContext(element)
      });

      candidates.push({
        key: normalise(attribute.name),
        rawKey: attribute.name,
        value: attribute.value,
        path,
        context: elementContext(element)
      });
    });
  });

  return candidates.filter((candidate) => String(candidate.value ?? "").trim() !== "");
}

function findXmlValue(candidates, mapRow) {
  const mappedKey = normalise(mapRow.mappedXml);
  const filters = [mapRow.dsp, mapRow.description]
    .filter(isRelevant)
    .map(normalise)
    .filter(Boolean);

  const matching = candidates
    .filter((candidate) => candidate.key === mappedKey || candidate.key.endsWith(mappedKey))
    .map((candidate) => {
      const score = filters.reduce((total, filter) => total + (candidate.context.includes(filter) ? 1 : 0), 0);
      return { ...candidate, score };
    })
    .filter((candidate) => filters.length === 0 || candidate.score === filters.length)
    .sort((a, b) => b.score - a.score || a.path.length - b.path.length);

  return matching[0] || null;
}

function formatForTarget(previousValue, xmlValue) {
  const raw = String(xmlValue ?? "").trim();
  const previous = String(previousValue ?? "").trim();
  if (!/^-?\d+$/.test(raw)) return raw;

  const decimal = Number(raw);
  if (!Number.isSafeInteger(decimal)) return raw;

  if (/^0x[0-9a-f]+$/i.test(previous)) {
    const width = Math.max(previous.length - 2, 2);
    return `0x${decimal.toString(16).toUpperCase().padStart(width, "0")}`;
  }

  if (/^[0-9a-f]+$/i.test(previous) && /[a-f]/i.test(previous)) {
    return decimal.toString(16).toUpperCase().padStart(previous.length, "0");
  }

  return raw;
}

function csvEscape(value) {
  const text = String(value ?? "");
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function buildLog(records, version, metadata = {}) {
  const lines = [
    `XML-XLS converter version,${csvEscape(version)}`,
    `Generated,${csvEscape(new Date().toISOString())}`,
    `XLS worksheet,${csvEscape(metadata.xlsSheetName || "")}`,
    `XLS header row,${metadata.xlsHeaderRow || ""}`,
    `XLS data start row,${metadata.xlsDataStartRow || ""}`,
    `MAP worksheet,${csvEscape(metadata.mapSheetName || "")}`,
    `MAP header row,${metadata.mapHeaderRow || ""}`,
    "",
    OUTPUT_HEADERS.map(csvEscape).join(",")
  ];

  records.forEach((record) => {
    lines.push([
      record.dsp,
      record.description,
      record.parameters,
      record.column,
      record.previousValue,
      record.newValue,
      record.mappedXml
    ].map(csvEscape).join(","));
  });

  return lines.join("\r\n");
}

function baseName(fileName) {
  return String(fileName || "converted").replace(/\.[^.]+$/, "");
}

function buildOutputNames(version) {
  const cleanedVersion = version.replace(/[\\/:*?"<>|]+/g, "_");
  const typedName = elements.outputNameInput.value.trim();
  const outputBase = typedName || `${baseName(state.xlsFile.name)}_${cleanedVersion}_converted`;
  return {
    workbookName: /\.xls[xm]?$/i.test(outputBase) ? outputBase : `${outputBase}.xlsx`,
    logName: `${outputBase}_log.csv`
  };
}

async function convert() {
  try {
    elements.convertButton.disabled = true;
    elements.downloadXlsButton.disabled = true;
    setProgress(0);
    setStatus("Converting files...", "", false);

    const version = elements.versionInput.value.trim();
    if (!version) throw new Error("Version number must not be empty.");
    if (!state.xmlFile) throw new Error("XML file is required.");
    if (!state.xlsFile) throw new Error("Input XLS file is required.");
    if (!state.mapFile) throw new Error("MAP file is required.");
    if (!window.XLSX) throw new Error("Spreadsheet library could not be loaded. Check the network connection and reload the page.");
    setProgress(10);

    const [xmlText, targetWorkbook, mapWorkbook] = await Promise.all([
      readFileAsText(state.xmlFile),
      readWorkbook(state.xlsFile),
      readWorkbook(state.mapFile)
    ]);
    setProgress(35);

    const xmlDocument = new DOMParser().parseFromString(xmlText, "application/xml");
    if (xmlDocument.querySelector("parsererror")) throw new Error("XML file could not be parsed.");
    setProgress(50);

    const candidates = buildXmlCandidates(xmlDocument);
    const mapData = extractRows(mapWorkbook, REQUIRED_MAP_HEADERS, {
      label: "MAP file",
      sheetName: TARGET_MAP_SHEET_NAME,
      allowSheetFallback: true
    });
    const targetData = extractRows(targetWorkbook, ["DSP", "Description", "Parameters"], {
      label: "XLS file",
      sheetName: TARGET_XLS_SHEET_NAME,
      allowSheetFallback: true,
      fixedHeaderRowNumbers: XLS_HEADER_ROW_CANDIDATES
    });
    setProgress(65);
    const targetIndex = buildTargetRowIndex(targetData.rows);
    const records = [];
    const misses = [];

    mapData.rows.filter((row) => isRelevant(row.mappedXml)).forEach((mapRow) => {
      const targetRow = targetIndex.get(rowKey(mapRow.dsp, mapRow.description, mapRow.parameters));
      const xmlMatch = findXmlValue(candidates, mapRow);

      if (!targetRow || !xmlMatch) {
        misses.push(`${mapRow.dsp} / ${mapRow.description} / ${mapRow.parameters} / ${mapRow.mappedXml}`);
        return;
      }

      TARGET_COLUMNS.forEach((columnNumber) => {
        const address = cellAddress(targetRow.rowIndex, columnNumber);
        const previousCell = targetData.sheet[address];
        const previousValue = previousCell?.v ?? "";
        const newValue = formatForTarget(previousValue, xmlMatch.value);

        if (String(previousValue) !== String(newValue)) {
          setSheetCell(targetData.sheet, targetRow.rowIndex, columnNumber, newValue, previousCell);
          records.push({
            dsp: targetRow.dsp,
            description: targetRow.description,
            parameters: targetRow.parameters,
            column: columnNumber,
            previousValue,
            newValue,
            mappedXml: mapRow.mappedXml
          });
        }
      });
    });
    setProgress(90);

    state.workbook = targetWorkbook;
    const names = buildOutputNames(version);
    state.outputName = names.workbookName;
    state.logName = names.logName;
    state.logText = buildLog(records, version, {
      xlsSheetName: targetData.sheetName,
      xlsHeaderRow: targetData.header.rowNumber,
      xlsDataStartRow: targetData.header.rowNumber + 1,
      mapSheetName: mapData.sheetName,
      mapHeaderRow: mapData.header.rowNumber
    });

    elements.logOutput.value = state.logText;
    elements.logMeta.textContent = `${records.length} modified cells, ${misses.length} unmapped rows`;
    elements.downloadXlsButton.disabled = false;
    setProgress(100);

    const summary = misses.length ? `${records.length} changes, ${misses.length} rows skipped` : `${records.length} changes`;
    const warnings = [targetData.warning, mapData.warning].filter(Boolean).join(" ");
    const debugDetails = `XLS worksheet: ${targetData.sheetName}, header row: ${targetData.header.rowNumber}. MAP worksheet: ${mapData.sheetName}, header row: ${mapData.header.rowNumber}.${warnings ? ` ${warnings}` : ""}`;
    setStatus("Conversion complete.", summary, false, debugDetails);
  } catch (error) {
    setProgress(0);
    setStatus(error.message, "", true);
  } finally {
    elements.convertButton.disabled = false;
  }
}

function downloadWorkbook() {
  if (!state.workbook) return;
  const outputName = state.outputName || "converted.xlsx";
  const bookType = /\.xls$/i.test(outputName) ? "xls" : "xlsx";
  XLSX.writeFile(state.workbook, outputName, { bookType });
}

function downloadLog() {
  if (!state.logText) return;
  const blob = new Blob([state.logText], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = state.logName || "conversion_log.csv";
  link.click();
  URL.revokeObjectURL(link.href);
}

wireDropZones();
elements.convertButton.addEventListener("click", convert);
elements.downloadXlsButton.addEventListener("click", downloadWorkbook);
elements.versionInput.addEventListener("input", () => {
  updateOutputNameSuggestion();
  setStatus("Waiting for files.", readinessSummary());
});
elements.outputNameInput.addEventListener("input", updateOutputNameSuggestion);
