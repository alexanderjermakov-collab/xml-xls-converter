const APP_VERSION = "1.1";
const RELEASE_DATE = "2026-05-22";
const APPLICATION_NAME = "Sharp Titan TV. AQ. XML-XLS converter";
const XLS_HEADER_ROW_NUMBER = 25;
const XLS_DATA_START_ROW_NUMBER = 26;
const MAP_HEADER_SCAN_LIMIT = 100;
const TARGET_XLS_SHEET_NAME = "AQ_Tbl";
const TARGET_MAP_SHEET_NAME = "MAP";
const DEFAULT_FILE_NAMES = {
  xml: "XML.xml",
  map: "MAP.xls",
  xls: "MAP.xls"
};
const REQUIRED_MAP_HEADERS = ["DSP", "Description", "Parameters", "Mapped to DAP XML"];
const OUTPUT_HEADERS = ["DSP", "Description", "Parameters", "TV Model", "Gain Column", "Previous Value", "New Value", "XML Mapping"];
const TV_TARGETS = {
  "24": {
    label: "1T-C24JF2x55E(K)B",
    modelColumn: 5,
    gainColumn: 6
  },
  "32": {
    label: "T-C32JF2x55E(K)B / 1T-C32JF3x55E(K)B / 2T-C32JF2x55E(K)B",
    modelColumn: 8,
    gainColumn: 9
  },
  "40": {
    label: "2T-C40JF2x55KE(K)B / 2T-C40JF3x55KE(K)B",
    modelColumn: 11,
    gainColumn: 12
  },
  "43": {
    label: "2T-C43JF2x55E(K)B",
    modelColumn: 14,
    gainColumn: 15
  }
};
const LINKED_SIZE_TARGETS = ["32", "40", "43"];

const state = {
  xmlFile: null,
  xlsFile: null,
  mapFile: null,
  workbook: null,
  outputName: "",
  outputBlob: null,
  logWorkbook: null,
  logText: "",
  logName: ""
};

const elements = {
  versionInput: document.getElementById("versionInput"),
  xmlInput: document.getElementById("xmlInput"),
  xlsInput: document.getElementById("xlsInput"),
  mapInput: document.getElementById("mapInput"),
  outputNameInput: document.getElementById("outputNameInput"),
  sameAqSettingInput: document.getElementById("sameAqSettingInput"),
  convertButton: document.getElementById("convertButton"),
  downloadXlsButton: document.getElementById("downloadXlsButton"),
  downloadLogButton: document.getElementById("downloadLogButton"),
  logOutput: document.getElementById("logOutput"),
  logMeta: document.getElementById("logMeta"),
  manualOutput: document.getElementById("manualOutput"),
  manualMeta: document.getElementById("manualMeta"),
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

function columnLetter(columnNumber) {
  return XLSX.utils.encode_col(columnNumber - 1);
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
    elements.versionInput.value.trim() ? "AQ version" : null,
    selectedTargetKeys().length ? "TV model" : null
  ].filter(Boolean);
  return ready.length ? `${ready.join(", ")} ready` : "";
}

function selectedModelKeys() {
  return Array.from(document.querySelectorAll('input[name="tvModel"]:checked')).map((input) => input.value);
}

function selectedTargetKeys() {
  const selected = selectedModelKeys();
  if (elements.sameAqSettingInput.checked && selected.some((key) => LINKED_SIZE_TARGETS.includes(key))) {
    return [...LINKED_SIZE_TARGETS];
  }
  return selected;
}

function selectedTargets() {
  return selectedTargetKeys().map((key) => TV_TARGETS[key]).filter(Boolean);
}

function updateSameAqOption() {
  const selected = selectedModelKeys();
  const canApplyLinked = selected.some((key) => LINKED_SIZE_TARGETS.includes(key));
  elements.sameAqSettingInput.disabled = !canApplyLinked;
  if (!canApplyLinked) elements.sameAqSettingInput.checked = false;
  setStatus("Waiting for files.", readinessSummary());
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
  return XLSX.read(buffer, { type: "array", cellDates: true, cellNF: true, cellStyles: true, bookVBA: true });
}

async function fetchDefaultFile(fileName, mimeType) {
  try {
    const response = await fetch(fileName);
    if (!response.ok) throw new Error();
    const blob = await response.blob();
    return new File([blob], fileName, { type: mimeType });
  } catch (_) {
    throw new Error(`Default file "${fileName}" was not found or cannot be loaded by the browser. Please select the file manually.`);
  }
}

async function resolveInputFiles() {
  const xmlFile = state.xmlFile || await fetchDefaultFile(DEFAULT_FILE_NAMES.xml, "text/xml");
  const xlsFile = state.xlsFile || await fetchDefaultFile(DEFAULT_FILE_NAMES.xls, "application/vnd.ms-excel");
  const mapFile = state.mapFile || await fetchDefaultFile(DEFAULT_FILE_NAMES.map, "application/vnd.ms-excel");
  return { xmlFile, xlsFile, mapFile };
}

async function readFormattingWorkbook(file) {
  if (!window.XlsxPopulate) {
    throw new Error("Formatting-preserving workbook writer could not be loaded. Check the network connection and reload the page.");
  }

  if (!/\.xlsm?x?$/i.test(file.name) || /\.xls$/i.test(file.name)) {
    throw new Error("Formatting-preserving conversion requires an .xlsx or .xlsm input file. Legacy .xls files cannot be safely rewritten in the browser without formatting loss.");
  }

  const buffer = await readFileAsArrayBuffer(file);
  return XlsxPopulate.fromDataAsync(buffer);
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
  const textValue = String(value);
  const existingCell = previousCell ? { ...previousCell } : { t: "s" };

  if (previousCell?.t === "n" && textValue.trim() !== "" && !Number.isNaN(Number(textValue))) {
    existingCell.v = Number(textValue);
    existingCell.t = "n";
  } else {
    existingCell.v = textValue;
    existingCell.t = "s";
  }

  delete existingCell.w;

  sheet[address] = existingCell;
}

function setPopulateCell(sheet, rowIndex, columnNumber, value, previousCell) {
  const textValue = String(value);
  const shouldBeNumber = previousCell?.t === "n" && textValue.trim() !== "" && !Number.isNaN(Number(textValue));
  sheet.cell(rowIndex + 1, columnNumber).value(shouldBeNumber ? Number(textValue) : textValue);
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
    `Application,${csvEscape(APPLICATION_NAME)}`,
    `Application version,${csvEscape(APP_VERSION)}`,
    `AQ settings version,${csvEscape(version)}`,
    `Released,${csvEscape(RELEASE_DATE)}`,
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
      record.tvModel,
      record.gainColumn,
      record.previousValue,
      record.newValue,
      record.mappedXml
    ].map(csvEscape).join(","));
  });

  return lines.join("\r\n");
}

function buildLogWorkbook(records, version, metadata = {}) {
  const rows = [
    ["Application", APPLICATION_NAME],
    ["Application version", APP_VERSION],
    ["AQ settings version", version],
    ["Generated", new Date().toISOString()],
    ["XLS worksheet", metadata.xlsSheetName || ""],
    ["XLS header row", metadata.xlsHeaderRow || ""],
    ["XLS data start row", metadata.xlsDataStartRow || ""],
    ["MAP worksheet", metadata.mapSheetName || ""],
    ["MAP header row", metadata.mapHeaderRow || ""],
    [],
    OUTPUT_HEADERS
  ];

  records.forEach((record) => {
    rows.push([
      record.dsp,
      record.description,
      record.parameters,
      record.tvModel,
      record.gainColumn,
      record.previousValue,
      record.newValue,
      record.mappedXml
    ]);
  });

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet["!cols"] = [
    { wch: 28 },
    { wch: 36 },
    { wch: 32 },
    { wch: 38 },
    { wch: 12 },
    { wch: 16 },
    { wch: 16 },
    { wch: 28 }
  ];
  XLSX.utils.book_append_sheet(workbook, sheet, "LOG");
  return workbook;
}

function buildManualOutput(records) {
  const byRow = new Map();
  records.forEach((record) => {
    if (!byRow.has(record.rowIndex)) {
      byRow.set(record.rowIndex, {
        dsp: record.dsp,
        description: record.description,
        parameters: record.parameters,
        values: []
      });
    }
    byRow.get(record.rowIndex).values.push(record.newValue);
  });

  return Array.from(byRow.values())
    .map((item) => item.values[0] ?? "")
    .join("\n");
}

function baseName(fileName) {
  return String(fileName || "converted").replace(/\.[^.]+$/, "");
}

function buildOutputNames(version, sourceFile = state.xlsFile) {
  const cleanedVersion = version.replace(/[\\/:*?"<>|]+/g, "_");
  const typedName = elements.outputNameInput.value.trim();
  const outputBase = typedName || `${baseName(sourceFile?.name || "AQ_settings")}_${cleanedVersion}_converted`;
  return {
    workbookName: /\.xlsm$/i.test(outputBase) ? outputBase : `${outputBase.replace(/\.xls$/i, "")}.xlsx`,
    logName: `${outputBase.replace(/\.(xlsx|xlsm|xls)$/i, "")}_log.xlsx`
  };
}

async function convert() {
  try {
    elements.convertButton.disabled = true;
    elements.downloadXlsButton.disabled = true;
    elements.downloadLogButton.disabled = true;
    state.outputBlob = null;
    state.logWorkbook = null;
    setProgress(0);
    setStatus("Converting files...", "", false);

    const version = elements.versionInput.value.trim() || "1.0";
    const targets = selectedTargets();
    if (!version) throw new Error("Version number must not be empty.");
    if (!targets.length) throw new Error("At least one TV model must be selected.");
    if (!window.XLSX) throw new Error("Spreadsheet library could not be loaded. Check the network connection and reload the page.");
    const inputFiles = await resolveInputFiles();
    setProgress(10);

    const [xmlText, targetWorkbook, mapWorkbook, formattingWorkbook] = await Promise.all([
      readFileAsText(inputFiles.xmlFile),
      readWorkbook(inputFiles.xlsFile),
      readWorkbook(inputFiles.mapFile),
      readFormattingWorkbook(inputFiles.xlsFile)
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
      fixedHeaderRowNumber: XLS_HEADER_ROW_NUMBER
    });
    const populateSheet = formattingWorkbook.sheet(targetData.sheetName);
    if (!populateSheet) throw new Error(`Formatting workbook worksheet "${targetData.sheetName}" was not found.`);
    populateSheet.cell("B2").value(version);
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

      targets.forEach((target) => {
        const columnNumber = target.gainColumn;
        const address = cellAddress(targetRow.rowIndex, columnNumber);
        const previousCell = targetData.sheet[address];
        const previousValue = previousCell?.v ?? "";
        const newValue = formatForTarget(previousValue, xmlMatch.value);

        if (String(previousValue) !== String(newValue)) {
          setSheetCell(targetData.sheet, targetRow.rowIndex, columnNumber, newValue, previousCell);
          setPopulateCell(populateSheet, targetRow.rowIndex, columnNumber, newValue, previousCell);
          records.push({
            rowIndex: targetRow.rowIndex,
            dsp: targetRow.dsp,
            description: targetRow.description,
            parameters: targetRow.parameters,
            tvModel: target.label,
            gainColumn: columnLetter(columnNumber),
            previousValue,
            newValue,
            mappedXml: mapRow.mappedXml
          });
        }
      });
    });
    setProgress(90);

    const outputBlob = await formattingWorkbook.outputAsync({ type: "blob" });
    state.workbook = targetWorkbook;
    state.outputBlob = outputBlob;
    const names = buildOutputNames(version, inputFiles.xlsFile);
    state.outputName = names.workbookName;
    state.logName = names.logName.replace(/\.csv$/i, ".xlsx");
    const logMetadata = {
      xlsSheetName: targetData.sheetName,
      xlsHeaderRow: targetData.header.rowNumber,
      xlsDataStartRow: XLS_DATA_START_ROW_NUMBER,
      mapSheetName: mapData.sheetName,
      mapHeaderRow: mapData.header.rowNumber
    };
    state.logText = buildLog(records, version, logMetadata);
    state.logWorkbook = buildLogWorkbook(records, version, logMetadata);

    elements.logOutput.value = state.logText;
    elements.logMeta.textContent = `${records.length} modified cells, ${misses.length} unmapped rows`;
    elements.manualOutput.value = buildManualOutput(records);
    elements.manualMeta.textContent = `${records.length ? new Set(records.map((record) => record.rowIndex)).size : 0} gain values generated for manual copy/paste`;
    elements.downloadXlsButton.disabled = false;
    elements.downloadLogButton.disabled = false;
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
  if (!state.outputBlob) return;
  const outputName = state.outputName || "converted.xlsx";
  const link = document.createElement("a");
  link.href = URL.createObjectURL(state.outputBlob);
  link.download = outputName;
  link.click();
  URL.revokeObjectURL(link.href);
}

function downloadLog() {
  if (!state.logWorkbook) return;
  XLSX.writeFile(state.logWorkbook, state.logName || "conversion_log.xlsx", { bookType: "xlsx" });
}

wireDropZones();
elements.convertButton.addEventListener("click", convert);
elements.downloadXlsButton.addEventListener("click", downloadWorkbook);
elements.downloadLogButton.addEventListener("click", downloadLog);
elements.versionInput.addEventListener("input", () => {
  updateOutputNameSuggestion();
  setStatus("Waiting for files.", readinessSummary());
});
elements.outputNameInput.addEventListener("input", updateOutputNameSuggestion);
document.querySelectorAll('input[name="tvModel"]').forEach((input) => input.addEventListener("change", updateSameAqOption));
elements.sameAqSettingInput.addEventListener("change", updateSameAqOption);
