(function () {
  const APP_NAME = 'Sharp Titan TV. AQ. XML-XLS converter';
  const APP_VERSION = '2.0';
  const RELEASE_DATE = '2026-05-25';
  const SUPPORTED_PROFILES = ['Movie', 'Music', 'Voice', 'User Selectable'];
  const EXCLUDED_PROFILES = ['Game', 'Night', 'Off'];
  const DEFAULT_OUTPUT_NAME = 'DAP_XML_CopyPaste_v2.xlsx';
  const LOG_OUTPUT_NAME = 'DAP_XML_CopyPaste_v2_LOG.xlsx';

  const state = {
    xmlFile: null,
    outputWorkbook: null,
    logWorkbook: null,
    outputName: DEFAULT_OUTPUT_NAME,
    logName: LOG_OUTPUT_NAME,
  };

  const xmlInput = document.getElementById('xmlInput');
  const xmlDropZone = document.getElementById('xmlDropZone');
  const xmlFileName = document.getElementById('xmlFileName');
  const outputFileName = document.getElementById('outputFileName');
  const convertButton = document.getElementById('convertButton');
  const downloadXlsxButton = document.getElementById('downloadXlsxButton');
  const downloadLogButton = document.getElementById('downloadLogButton');
  const progressBar = document.getElementById('progressBar');
  const statusWindow = document.querySelector('.status-window');
  const statusText = document.getElementById('statusText');
  const debugText = document.getElementById('debugText');
  const summaryBody = document.getElementById('summaryBody');

  function setProgress(percent) {
    progressBar.style.width = `${Math.max(0, Math.min(100, percent))}%`;
  }

  function setStatus(message, detail, type) {
    statusWindow.classList.remove('status-error', 'status-warning', 'status-ok');
    if (type) {
      statusWindow.classList.add(`status-${type}`);
    }
    statusText.textContent = message;
    debugText.textContent = detail || '';
  }

  function normalizeName(value) {
    return String(value || '')
      .trim()
      .replace(/^<+/, '')
      .replace(/>+$/, '')
      .toLowerCase();
  }

  function localName(element) {
    return (element.localName || element.nodeName || '').toLowerCase();
  }

  function attr(element, name) {
    return element.getAttribute(name) || '';
  }

  function readNodeValue(element) {
    if (!element) {
      return '';
    }
    const preferredAttributes = ['value', 'default', 'enabled'];
    for (const name of preferredAttributes) {
      const value = element.getAttribute(name);
      if (value !== null && String(value).trim() !== '') {
        return String(value).trim();
      }
    }
    return String(element.textContent || '').trim();
  }

  function formatHex(value) {
    if (value === null || value === undefined || value === '') {
      return '';
    }
    const text = String(value).trim();
    if (!/^[+-]?\d+$/.test(text)) {
      return '';
    }
    let number = BigInt(text);
    const modulo = 0x100000000n;
    if (number < 0) {
      number = modulo + (number % modulo);
    }
    number %= modulo;
    return number.toString(16).toUpperCase().padStart(8, '0');
  }

  function parseNumericList(value) {
    const parts = String(value || '')
      .split(/[,\s]+/)
      .map((part) => part.trim())
      .filter(Boolean);
    if (!parts.length || parts.some((part) => !/^[+-]?\d+$/.test(part))) {
      return null;
    }
    return parts;
  }

  function findInternalSpeakerEndpoint(xmlDoc) {
    const endpoints = Array.from(xmlDoc.getElementsByTagName('*')).filter((node) => localName(node) === 'endpoint');
    const internal = endpoints.find((node) => normalizeName(attr(node, 'type')) === 'internal_speaker');
    if (internal) {
      return internal;
    }
    return endpoints.find((node) => normalizeName(attr(node, 'type')) !== 'headphones') || null;
  }

  function getProfiles(endpoint) {
    const profileMap = new Map();
    const profiles = Array.from(endpoint.getElementsByTagName('*')).filter((node) => localName(node) === 'profile');
    profiles.forEach((node) => {
      const name = attr(node, 'name') || attr(node, 'type');
      if (SUPPORTED_PROFILES.includes(name) && !profileMap.has(name)) {
        profileMap.set(name, node);
      }
    });
    return profileMap;
  }

  function findParameter(profileNode, parameterName) {
    const normalized = normalizeName(parameterName);
    if (!normalized) {
      return null;
    }
    return Array.from(profileNode.getElementsByTagName('*')).find((node) => localName(node) === normalized) || null;
  }

  function computeMappedRow(mapRow, profileMap, endpointNode, logRows, profileStats) {
    const profileName = mapRow.xmlProfile || '';
    const parameterName = mapRow.xmlParameter || '';
    const normalizedParameter = normalizeName(parameterName);
    const profileNode = profileMap.get(profileName);
    const result = {
      profile: profileName,
      xmlParameter: parameterName,
      xmlValueDec: mapRow.xmlValueDec || '',
      dsp: mapRow.dsp || '',
      description: mapRow.description || '',
      parameters: mapRow.parameters || '',
      gainHex: mapRow.gainHex || '',
      status: 'Default from MAP',
    };

    if (!profileStats[profileName]) {
      profileStats[profileName] = { mapped: 0, missing: 0, defaulted: 0 };
    }

    if (!profileNode) {
      profileStats[profileName].missing += 1;
      result.status = 'Profile not found';
      logRows.push([profileName, parameterName, result.parameters, '', result.gainHex, result.status]);
      return result;
    }

    if (!normalizedParameter || normalizedParameter === 'missing') {
      profileStats[profileName].defaulted += 1;
      logRows.push([profileName, parameterName, result.parameters, '', result.gainHex, result.status]);
      return result;
    }

    const parameterNode = findParameter(profileNode, normalizedParameter) || findParameter(endpointNode, normalizedParameter);
    if (!parameterNode) {
      profileStats[profileName].missing += 1;
      result.status = 'XML parameter not found';
      logRows.push([profileName, parameterName, result.parameters, '', result.gainHex, result.status]);
      return result;
    }

    const rawValue = readNodeValue(parameterNode);
    if (!rawValue) {
      profileStats[profileName].defaulted += 1;
      result.status = 'Container found; MAP default kept';
      logRows.push([profileName, parameterName, result.parameters, '', result.gainHex, result.status]);
      return result;
    }

    let valueForHex = rawValue;
    const list = parseNumericList(rawValue);

    if (list && /nb[_\s-]*bands/i.test(result.parameters)) {
      valueForHex = String(list.length);
    } else if (list && list.length > 1) {
      profileStats[profileName].defaulted += 1;
      result.xmlValueDec = rawValue;
      result.status = 'List value found; MAP default kept';
      logRows.push([profileName, parameterName, result.parameters, rawValue, result.gainHex, result.status]);
      return result;
    }

    const hexValue = formatHex(valueForHex);
    if (!hexValue) {
      profileStats[profileName].defaulted += 1;
      result.xmlValueDec = rawValue;
      result.status = 'Non-decimal value found; MAP default kept';
      logRows.push([profileName, parameterName, result.parameters, rawValue, result.gainHex, result.status]);
      return result;
    }

    profileStats[profileName].mapped += 1;
    result.xmlValueDec = valueForHex;
    result.gainHex = hexValue;
    result.status = 'Mapped from XML';
    logRows.push([profileName, parameterName, result.parameters, valueForHex, hexValue, result.status]);
    return result;
  }

  function buildOutputWorkbook(mappedRows, inputName, outputName) {
    const now = new Date();
    const aoa = [
      [APP_NAME],
      ['Version', APP_VERSION],
      ['Release date', RELEASE_DATE],
      ['Conversion date', now.toLocaleString()],
      ['Input XML file', inputName],
      ['Output XLSX file', outputName],
      ['User instructions', 'Open this workbook and manually Copy/Paste the generated table data into the TitanOS AQ settings workbook.'],
      [],
      ['DAP Audio settings in XML file', '', '', 'DAP Audio settings in XLS file', '', '', ''],
      ['XML Profile name', 'XML parameter name', 'XML parameter value (DEC)', 'DSP', 'Description', 'Parameters', 'Gain (Hex)'],
    ];

    mappedRows.forEach((row) => {
      aoa.push([
        row.profile,
        row.xmlParameter,
        row.xmlValueDec,
        row.dsp,
        row.description,
        row.parameters,
        row.gainHex,
      ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [
      { wch: 22 },
      { wch: 34 },
      { wch: 24 },
      { wch: 30 },
      { wch: 42 },
      { wch: 34 },
      { wch: 14 },
    ];
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } },
      { s: { r: 8, c: 0 }, e: { r: 8, c: 2 } },
      { s: { r: 8, c: 3 }, e: { r: 8, c: 6 } },
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Copy&Paste');
    return wb;
  }

  function buildLogWorkbook(logRows, summary) {
    const aoa = [
      [APP_NAME],
      ['Version', APP_VERSION],
      ['Conversion date', new Date().toLocaleString()],
      ['Internal endpoint', summary.endpointType],
      ['Profiles used', SUPPORTED_PROFILES.join(', ')],
      ['Profiles excluded', EXCLUDED_PROFILES.join(', ')],
      ['Total MAP rows', summary.totalRows],
      ['Mapped rows', summary.mappedRows],
      ['Missing XML/profile rows', summary.missingRows],
      ['Default rows kept', summary.defaultRows],
      [],
      ['Profile', 'XML parameter', 'XLS Parameters', 'XML value DEC', 'Gain HEX', 'Operation'],
      ...logRows,
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [
      { wch: 20 },
      { wch: 34 },
      { wch: 34 },
      { wch: 24 },
      { wch: 14 },
      { wch: 32 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'LOG');
    return wb;
  }

  function renderSummary(profileStats) {
    summaryBody.innerHTML = '';
    SUPPORTED_PROFILES.forEach((profile) => {
      const stats = profileStats[profile] || { mapped: 0, missing: 0, defaulted: 0 };
      const tr = document.createElement('tr');
      [profile, stats.mapped, stats.missing, stats.defaulted].forEach((value) => {
        const td = document.createElement('td');
        td.textContent = value;
        tr.appendChild(td);
      });
      summaryBody.appendChild(tr);
    });
  }

  function makeOutputName(inputName) {
    const baseName = inputName.replace(/\.[^.]+$/, '') || 'DAP_XML';
    return `${baseName}_CopyPaste_v2.xlsx`;
  }

  async function convert() {
    if (!state.xmlFile) {
      setStatus('XML file is required.', 'Select the Dolby Tuning Tool XML export.', 'error');
      return;
    }

    if (!window.XLSX) {
      setStatus('XLSX library is not loaded.', 'Check internet access and reload the page.', 'error');
      return;
    }

    setProgress(12);
    setStatus('Reading XML file...', state.xmlFile.name);

    try {
      const xmlText = await state.xmlFile.text();
      setProgress(28);

      const xmlDoc = new DOMParser().parseFromString(xmlText, 'application/xml');
      const parserError = xmlDoc.querySelector('parsererror');
      if (parserError) {
        throw new Error('XML parser error. Please check that the input file is valid XML.');
      }

      const endpoint = findInternalSpeakerEndpoint(xmlDoc);
      if (!endpoint) {
        throw new Error('Endpoint type "internal_speaker" was not found in XML file.');
      }

      const profileMap = getProfiles(endpoint);
      setProgress(48);

      const logRows = [];
      const profileStats = {};
      const mappedRows = window.MAP20_ROWS.map((row) => computeMappedRow(row, profileMap, endpoint, logRows, profileStats));
      setProgress(72);

      const totalMapped = Object.values(profileStats).reduce((sum, item) => sum + item.mapped, 0);
      const totalMissing = Object.values(profileStats).reduce((sum, item) => sum + item.missing, 0);
      const totalDefault = Object.values(profileStats).reduce((sum, item) => sum + item.defaulted, 0);
      const summary = {
        endpointType: attr(endpoint, 'type') || 'internal_speaker',
        totalRows: mappedRows.length,
        mappedRows: totalMapped,
        missingRows: totalMissing,
        defaultRows: totalDefault,
      };

      state.outputName = makeOutputName(state.xmlFile.name);
      state.logName = state.outputName.replace(/\.xlsx$/i, '_LOG.xlsx');
      state.outputWorkbook = buildOutputWorkbook(mappedRows, state.xmlFile.name, state.outputName);
      state.logWorkbook = buildLogWorkbook(logRows, summary);

      outputFileName.textContent = state.outputName;
      renderSummary(profileStats);
      downloadXlsxButton.disabled = false;
      downloadLogButton.disabled = false;
      setProgress(100);

      const message = totalMissing > 0 ? 'Conversion completed with warnings.' : 'Conversion completed successfully.';
      const type = totalMissing > 0 ? 'warning' : 'ok';
      setStatus(
        message,
        `${totalMapped} rows mapped from XML, ${totalDefault} rows kept from MAP defaults, ${totalMissing} rows missing XML/profile data.`,
        type
      );
    } catch (error) {
      state.outputWorkbook = null;
      state.logWorkbook = null;
      downloadXlsxButton.disabled = true;
      downloadLogButton.disabled = true;
      setProgress(0);
      setStatus('Conversion failed.', error.message, 'error');
    }
  }

  function downloadWorkbook(workbook, filename) {
    if (!workbook) {
      return;
    }
    XLSX.writeFile(workbook, filename, { bookType: 'xlsx' });
  }

  function handleFile(file) {
    state.xmlFile = file || null;
    state.outputWorkbook = null;
    state.logWorkbook = null;
    downloadXlsxButton.disabled = true;
    downloadLogButton.disabled = true;
    convertButton.disabled = !state.xmlFile;
    outputFileName.textContent = 'Output file will be created after conversion';
    setProgress(0);

    if (state.xmlFile) {
      xmlFileName.textContent = state.xmlFile.name;
      setStatus('XML file selected.', 'Press Convert to create the XLSX and LOG files.', 'ok');
    } else {
      xmlFileName.textContent = 'No file selected';
      setStatus('Select XML file to start.', 'Only the internal_speaker endpoint is used. Game, Night, Off and headphones data are excluded.');
    }
  }

  xmlInput.addEventListener('change', (event) => {
    handleFile(event.target.files[0]);
  });

  ['dragenter', 'dragover'].forEach((eventName) => {
    xmlDropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      xmlDropZone.classList.add('drag-over');
    });
  });

  ['dragleave', 'drop'].forEach((eventName) => {
    xmlDropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      xmlDropZone.classList.remove('drag-over');
    });
  });

  xmlDropZone.addEventListener('drop', (event) => {
    handleFile(event.dataTransfer.files[0]);
  });

  convertButton.addEventListener('click', convert);
  downloadXlsxButton.addEventListener('click', () => downloadWorkbook(state.outputWorkbook, state.outputName));
  downloadLogButton.addEventListener('click', () => downloadWorkbook(state.logWorkbook, state.logName));
})();
