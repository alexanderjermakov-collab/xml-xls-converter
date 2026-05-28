(function () {
  const APP_NAME = 'Sharp Titan TV. AQ. XML-XLS converter';
  const APP_VERSION = '2.6';
  const RELEASE_DATE = '2026-05-28';
  const SUPPORTED_PROFILES = ['Movie', 'Music', 'Voice', 'User Selectable'];
  const EXCLUDED_PROFILES = ['Game', 'Night', 'Off'];
  const DEFAULT_OUTPUT_NAME = 'Titan TV. DAP AQ. XML-XLS converter. Output.xlsx';
  const LOG_OUTPUT_NAME = 'Titan TV. DAP AQ. XML-XLS converter. Output LOG.xlsx';
  const EQ_CENTER_FREQUENCIES = [
    47, 141, 234, 328, 469, 656, 844, 1031, 1313, 1688,
    2250, 3000, 3750, 4688, 5813, 7125, 9000, 11250, 13875, 19688,
  ];

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

  function comparableName(value) {
    return normalizeName(value).replace(/[\s_\-./\\()[\]{}:;]+/g, '');
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

  function directText(element) {
    return Array.from(element.childNodes || [])
      .filter((node) => node.nodeType === Node.TEXT_NODE)
      .map((node) => String(node.nodeValue || '').trim())
      .filter(Boolean)
      .join(' ')
      .trim();
  }

  function nearestProfile(element) {
    let current = element.parentElement;
    while (current) {
      if (localName(current) === 'profile') {
        return current;
      }
      current = current.parentElement;
    }
    return null;
  }

  function parameterMatches(element, parameterName) {
    const expected = comparableName(parameterName);
    if (!expected) {
      return false;
    }

    if (comparableName(localName(element)) === expected) {
      return true;
    }

    const matchingAttributes = ['name', 'id', 'key', 'type', 'parameter', 'param'];
    return matchingAttributes.some((name) => comparableName(element.getAttribute(name)) === expected);
  }

  function parameterValue(element) {
    if (!element) {
      return '';
    }

    const preferredAttributes = ['value', 'val', 'default', 'enabled', 'hex', 'gain', 'data'];
    for (const name of preferredAttributes) {
      const value = element.getAttribute(name);
      if (value !== null && String(value).trim() !== '') {
        return String(value).trim();
      }
    }
    return directText(element) || readNodeValue(element);
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

  function hexToSignedDec(hexValue) {
    const text = String(hexValue || '').trim();
    if (!/^[0-9a-fA-F]{1,8}$/.test(text)) {
      return '';
    }
    let number = parseInt(text, 16);
    if (number > 0x7fffffff) {
      number -= 0x100000000;
    }
    return String(number);
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

  function bandNumber(parameters, suffix) {
    const match = String(parameters || '').match(/^Band_(\d{2})_(Fc|Target)$/i);
    if (!match || match[2].toLowerCase() !== suffix.toLowerCase()) {
      return 0;
    }
    return Number(match[1]);
  }

  function isEqBandListRow(row) {
    return /eq[_\s-]*nb[_\s-]*bands/i.test(row.parameters || '') && /bands?/i.test(row.xmlParameter || '');
  }

  function fixedDecForParameter(row) {
    const parameter = comparableName(row.parameters);

    if (parameter === 'deamountmax') {
      return '163';
    }
    if (parameter === 'deamountmin') {
      return '0';
    }
    if (parameter === 'surroundboostmax') {
      return '964';
    }
    if (parameter === 'surroundboostmin') {
      return '0';
    }
    if (parameter === 'bassboostmax' || parameter === 'bassboostmin') {
      return hexToSignedDec(row.gainHex);
    }

    return '';
  }

  function inferredXmlParameter(row) {
    const parameter = comparableName(row.parameters);
    if (parameter === 'speakerangle') {
      return 'virtualizer-surround-speaker-angle';
    }
    return normalizeName(row.xmlParameter);
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
      const profileType = attr(node, 'type');
      const profileName = attr(node, 'name');
      const supportedName = SUPPORTED_PROFILES.find((name) => {
        const supported = comparableName(name);
        return comparableName(profileType) === supported || comparableName(profileName) === supported;
      });

      if (supportedName && !profileMap.has(supportedName)) {
        profileMap.set(supportedName, node);
      }
    });
    return profileMap;
  }

  function findParameter(scopeNode, parameterName, options) {
    const nodes = Array.from(scopeNode.getElementsByTagName('*'));
    return nodes.find((node) => {
      if (options && options.excludeProfileChildren && nearestProfile(node)) {
        return false;
      }
      return parameterMatches(node, parameterName);
    }) || null;
  }

  function computeMappedRow(mapRow, profileMap, endpointNode, logRows, profileStats, activeEqParameter) {
    const profileName = mapRow.xmlProfile || '';
    const parameterName = mapRow.xmlParameter || '';
    const normalizedParameter = inferredXmlParameter(mapRow);
    const profileNode = profileMap.get(profileName);
    const bandFc = bandNumber(mapRow.parameters, 'Fc');
    const bandTarget = bandNumber(mapRow.parameters, 'Target');
    const result = {
      profile: profileName,
      xmlParameter: parameterName || normalizedParameter,
      xmlValueDec: mapRow.xmlValueDec || '',
      dsp: mapRow.dsp || '',
      description: mapRow.description || '',
      parameters: mapRow.parameters || '',
      gainHex: mapRow.gainHex || '',
      gainDec: hexToSignedDec(mapRow.gainHex),
      status: 'Default from MAP',
    };

    if (!profileStats[profileName]) {
      profileStats[profileName] = { mapped: 0, missing: 0, defaulted: 0 };
    }

    if (!profileNode) {
      profileStats[profileName].missing += 1;
      result.status = 'Profile not found';
      logRows.push([profileName, parameterName, result.parameters, '', result.gainDec, result.gainHex, result.status]);
      return result;
    }

    if (bandFc) {
      const decValue = EQ_CENTER_FREQUENCIES[bandFc - 1];
      if (decValue !== undefined) {
        profileStats[profileName].mapped += 1;
        result.xmlParameter = 'Fixed EQ center frequency';
        result.xmlValueDec = String(decValue);
        result.gainDec = String(decValue);
        result.gainHex = formatHex(decValue);
        result.status = 'Fixed EQ center frequency';
        logRows.push([profileName, result.xmlParameter, result.parameters, result.xmlValueDec, result.gainDec, result.gainHex, result.status]);
        return result;
      }
    }

    const fixedDecValue = fixedDecForParameter(mapRow);
    if (fixedDecValue !== '') {
      profileStats[profileName].mapped += 1;
      result.xmlValueDec = fixedDecValue;
      result.gainDec = fixedDecValue;
      result.gainHex = formatHex(fixedDecValue);
      result.status = 'Fixed min/max value';
      logRows.push([profileName, parameterName, result.parameters, fixedDecValue, result.gainDec, result.gainHex, result.status]);
      return result;
    }

    if (!normalizedParameter || normalizedParameter === 'missing') {
      if (bandTarget && activeEqParameter) {
        const activeNode =
          findParameter(profileNode, activeEqParameter) ||
          findParameter(endpointNode, activeEqParameter, { excludeProfileChildren: true });
        const activeList = parseNumericList(parameterValue(activeNode));
        const targetValue = activeList ? activeList[bandTarget - 1] : '';

        if (targetValue !== undefined && targetValue !== '') {
          const hexValue = formatHex(targetValue);
          profileStats[profileName].mapped += 1;
          result.xmlParameter = activeEqParameter;
          result.xmlValueDec = targetValue;
          result.gainDec = targetValue;
          result.gainHex = hexValue;
          result.status = 'EQ band target from XML list';
          logRows.push([profileName, result.xmlParameter, result.parameters, targetValue, result.gainDec, hexValue, result.status]);
          return result;
        }

        profileStats[profileName].missing += 1;
        result.xmlParameter = activeEqParameter;
        result.status = 'EQ band target value not found';
        logRows.push([profileName, result.xmlParameter, result.parameters, '', result.gainDec, result.gainHex, result.status]);
        return result;
      }

      profileStats[profileName].defaulted += 1;
      logRows.push([profileName, parameterName, result.parameters, '', result.gainDec, result.gainHex, result.status]);
      return result;
    }

    const parameterNode =
      findParameter(profileNode, normalizedParameter) ||
      findParameter(endpointNode, normalizedParameter, { excludeProfileChildren: true });
    if (!parameterNode) {
      profileStats[profileName].missing += 1;
      result.status = 'XML parameter not found';
      logRows.push([profileName, parameterName, result.parameters, '', result.gainDec, result.gainHex, result.status]);
      return result;
    }

    const rawValue = parameterValue(parameterNode);
    if (!rawValue) {
      profileStats[profileName].defaulted += 1;
      result.status = 'Container found; MAP default kept';
      logRows.push([profileName, parameterName, result.parameters, '', result.gainDec, result.gainHex, result.status]);
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
      logRows.push([profileName, parameterName, result.parameters, rawValue, result.gainDec, result.gainHex, result.status]);
      return result;
    }

    const hexValue = formatHex(valueForHex);
    if (!hexValue) {
      profileStats[profileName].defaulted += 1;
      result.xmlValueDec = rawValue;
      result.status = 'Non-decimal value found; MAP default kept';
      logRows.push([profileName, parameterName, result.parameters, rawValue, result.gainDec, result.gainHex, result.status]);
      return result;
    }

    profileStats[profileName].mapped += 1;
    result.xmlValueDec = valueForHex;
    result.gainDec = valueForHex;
    result.gainHex = hexValue;
    result.status = 'Mapped from XML';
    logRows.push([profileName, parameterName, result.parameters, valueForHex, result.gainDec, hexValue, result.status]);
    return result;
  }

  function mapRowsWithEqContext(profileMap, endpoint, logRows, profileStats) {
    let activeEqParameter = '';

    return window.MAP20_ROWS.map((row) => {
      if (isEqBandListRow(row)) {
        activeEqParameter = normalizeName(row.xmlParameter);
      }

      return computeMappedRow(row, profileMap, endpoint, logRows, profileStats, activeEqParameter);
    });
  }

  function buildOutputWorkbook(mappedRows, inputName, outputName) {
    const now = new Date();
    const instructionRows = [
      ['User Instruction', 'Use the manual Copy&Paste operation by following steps:'],
      ['Step 1', 'Use the DEC column and copy its contents (656 cells) into AQ.xls, starting from row 1166 (Entertainment Custom Mode), into the Gain (Dec) column corresponding to the currently tuned TV model.'],
    ];
    const aoa = [
      [APP_NAME],
      ['Version', APP_VERSION],
      ['Release date', RELEASE_DATE],
      ['Conversion date', now.toLocaleString()],
      ['Input XML file', inputName],
      ['Output XLSX file', outputName],
      [],
      ...instructionRows,
      [],
      ['DAP Audio settings in XML file', '', '', 'DAP Audio settings in XLS file', '', '', ''],
      ['XML Profile name', 'XML parameter name', 'XML parameter value (DEC)', 'DSP', 'Description', 'Parameters', 'Gain (Dec)'],
    ];

    mappedRows.forEach((row) => {
      aoa.push([
        row.profile,
        row.xmlParameter,
        row.xmlValueDec,
        row.dsp,
        row.description,
        row.parameters,
        row.gainDec,
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
      { s: { r: 10, c: 0 }, e: { r: 10, c: 2 } },
      { s: { r: 10, c: 3 }, e: { r: 10, c: 6 } },
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
      ['Profile', 'XML parameter', 'XLS Parameters', 'XML value DEC', 'Gain DEC', 'Gain HEX', 'Operation'],
      ...logRows,
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [
      { wch: 20 },
      { wch: 34 },
      { wch: 34 },
      { wch: 24 },
      { wch: 14 },
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

  function makeOutputName() {
    return DEFAULT_OUTPUT_NAME;
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
      const mappedRows = mapRowsWithEqContext(profileMap, endpoint, logRows, profileStats);
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

      state.outputName = makeOutputName();
      state.logName = LOG_OUTPUT_NAME;
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
