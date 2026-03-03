/**
 * AgriData Hub - CSV Logic
 */

document.addEventListener('DOMContentLoaded', () => {
    const csvInput = document.getElementById('csv-input');
    const dropZone = document.getElementById('drop-zone');
    const dashboard = document.getElementById('dashboard');
    const tableHead = document.getElementById('table-head');
    const tableBody = document.getElementById('table-body');
    const summaryCards = document.getElementById('summary-cards');
    const rowCountBadge = document.getElementById('row-count');
    const tableSearch = document.getElementById('table-search');
    const insuranceFilter = document.getElementById('insurance-filter');
    const classFilter = document.getElementById('class-filter');
    const cropFilter = document.getElementById('crop-filter');
    const dateFromFilter = document.getElementById('date-from-filter');
    const dateToFilter = document.getElementById('date-to-filter');
    const resetFiltersBtn = document.getElementById('reset-filters-btn');
    const printBtn = document.getElementById('print-btn');
    const clearDataBtn = document.getElementById('clear-data-btn');
    const reportTypeSelector = document.getElementById('report-type-selector');

    let csvData = [];
    let headers = [];
    let selectedRowIds = [];

    function calculateAge(birthdate) {
        if (!birthdate) return '';
        const birthDate = new Date(birthdate);
        if (isNaN(birthDate.getTime())) return '';
        const today = new Date();
        let age = today.getFullYear() - birthDate.getFullYear();
        const m = today.getMonth() - birthDate.getMonth();
        if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) {
            age--;
        }
        return age;
    }

    // Initialize IDB Storage
    loadFromStorage();

    // Make empty activity table cells editable for UI functionality
    document.querySelectorAll('.hvc-activity-table tbody td:not(:first-child)').forEach(td => {
        if (td.id && td.id.endsWith('-date')) return; // Skip auto-filled date cells
        td.setAttribute('contenteditable', 'true');
        td.style.outline = 'none';
        td.style.cursor = 'text';
    });

    // Drag and Drop Handlers
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        handleFiles(e.dataTransfer.files);
    });

    csvInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
    });

    clearDataBtn.addEventListener('click', () => {
        if (confirm('Are you sure you want to clear all data? This cannot be undone.')) {
            clearStorage();
        }
    });

    function handleFiles(files) {
        if (!files.length) return;

        const fileList = Array.from(files).filter(file => file.type === 'text/csv' || file.name.endsWith('.csv'));

        if (fileList.length === 0) {
            alert('Please upload valid CSV files.');
            return;
        }

        processFiles(fileList);
    }

    async function processFiles(files) {
        let processedCount = 0;
        let newRecords = [];
        let newHeaders = [];

        // Show loading state if needed
        dropZone.innerHTML = `<div class="upload-content"><h3>Processing ${files.length} files...</h3></div>`;

        // Calculate starting index based on existing data
        let currentIndex = csvData.length + 1;

        for (const file of files) {
            await new Promise((resolve) => {
                Papa.parse(file, {
                    header: true,
                    dynamicTyping: true,
                    skipEmptyLines: true,
                    complete: function (results) {
                        const rawData = results.data;
                        const rawHeaders = results.meta.fields;

                        if (!headers.length) headers = rawHeaders; // Set headers from first file if empty
                        // Ideally we should check if headers match, but for now we assume they are similar or at least we stick to the first file's structure.

                        // Normalization Logic
                        const normalize = (h) => h.toLowerCase().trim().replace(/[^a-z0-9]/g, '');
                        const headerMap = {};
                        rawHeaders.forEach(h => {
                            const norm = normalize(h);
                            if (norm === 'farmersid') headerMap['FarmersID'] = h;
                            else if (norm === 'farmid') headerMap['FarmID'] = h;
                            else if (norm === 'area') headerMap['Area'] = h;
                            else if (norm === 'amountcover') headerMap['AmountCover'] = h;
                            else if (norm === 'animaltype') headerMap['AnimalType'] = h;
                            else if (norm === 'boattype') headerMap['BoatType'] = h;
                            else if (norm === 'boatmaterial') headerMap['BoatMaterial'] = h;
                            else if (norm === 'insuranceline') headerMap['InsuranceLine'] = h;
                            else if (norm === 'rsbsaid') headerMap['RSBSAID'] = h;
                            else if (norm === 'lastname') headerMap['LastName'] = h;
                            else if (norm === 'firstname') headerMap['FirstName'] = h;
                            else if (norm === 'middlname') headerMap['MiddlName'] = h;
                            else if (norm === 'extname') headerMap['ExtName'] = h;
                            else if (norm === 'munfarmer') headerMap['MunFarmer'] = h;
                            else if (norm === 'brgyfarmer') headerMap['BrgyFarmer'] = h;
                            else if (norm === 'stfarmer') headerMap['StFarmer'] = h;
                            else if (norm === 'croptype') headerMap['CropType'] = h;
                            else if (norm === 'farmershare') headerMap['FarmerShare'] = h;
                            else if (norm === 'governmentshare') headerMap['GovernmentShare'] = h;
                            else if (norm === 'lender' || norm === 'lendername') headerMap['Lender'] = h;
                            else if (norm === 'groupname' || norm === 'group') headerMap['GroupName'] = h;
                            else if (norm === 'birthday' || norm === 'birthdate') headerMap['Birthdate'] = h;
                            else if (norm === 'relationship' || norm === 'benerelationship') headerMap['BeneRelationship'] = h;
                            else if (norm === 'beneficiary') headerMap['Beneficiary'] = h;
                            else if (norm === 'premium') headerMap['Premium'] = h;
                            else if (norm === 'planting' || norm === 'dateplanted' || norm === 'plantingdate') headerMap['Planting'] = h;
                        });

                        const transformedData = rawData.map(row => {
                            // Clone original row to preserve all background CSV data
                            const newRow = { ...row };

                            // Add Row Number
                            newRow['No.'] = currentIndex++;

                            if (headerMap['FarmersID']) newRow.FarmersID = row[headerMap['FarmersID']];
                            if (headerMap['FarmID']) newRow.FarmID = row[headerMap['FarmID']];
                            else newRow.FarmID = row['FarmID'] || ''; // Fallback

                            if (headerMap['Area']) newRow.Area = row[headerMap['Area']];
                            if (headerMap['AmountCover']) newRow.AmountCover = row[headerMap['AmountCover']];
                            if (headerMap['AnimalType']) newRow.AnimalType = row[headerMap['AnimalType']];
                            if (headerMap['BoatType']) newRow.BoatType = row[headerMap['BoatType']];
                            if (headerMap['BoatMaterial']) newRow.BoatMaterial = row[headerMap['BoatMaterial']];
                            if (headerMap['InsuranceLine']) newRow.InsuranceLine = row[headerMap['InsuranceLine']];

                            if (headerMap['RSBSAID']) newRow.RSBSAID = row[headerMap['RSBSAID']];
                            if (headerMap['LastName']) newRow.LastName = row[headerMap['LastName']];
                            if (headerMap['FirstName']) newRow.FirstName = row[headerMap['FirstName']];
                            if (headerMap['MiddlName']) newRow.MiddlName = row[headerMap['MiddlName']];
                            if (headerMap['ExtName']) newRow.ExtName = row[headerMap['ExtName']];
                            if (headerMap['Lender']) newRow.Lender = row[headerMap['Lender']];
                            if (headerMap['GroupName']) newRow.GroupName = row[headerMap['GroupName']];

                            // Concatenate Farmer Name
                            const lName = newRow.LastName || '';
                            const fName = newRow.FirstName || '';
                            const mName = newRow.MiddlName || '';
                            const extName = newRow.ExtName || '';

                            // Build full name: LastName, FirstName MiddlName ExtName
                            let fullName = `${lName}`;
                            if (fullName && (fName || mName || extName)) fullName += ', ';

                            let givenNames = [fName, mName, extName].filter(n => n.trim() !== '').join(' ');
                            fullName += givenNames;

                            newRow['Farmer Name'] = fullName.trim();

                            if (headerMap['MunFarmer']) newRow.MunFarmer = row[headerMap['MunFarmer']];
                            if (headerMap['BrgyFarmer']) newRow.BrgyFarmer = row[headerMap['BrgyFarmer']];
                            if (headerMap['StFarmer']) newRow.StFarmer = row[headerMap['StFarmer']];
                            if (headerMap['Birthdate']) newRow.Birthdate = row[headerMap['Birthdate']];
                            if (headerMap['BeneRelationship']) newRow.BeneRelationship = row[headerMap['BeneRelationship']];
                            if (headerMap['Beneficiary']) newRow.Beneficiary = row[headerMap['Beneficiary']];
                            if (headerMap['Premium']) newRow.Premium = row[headerMap['Premium']];
                            if (headerMap['CropType']) newRow.CropType = row[headerMap['CropType']];
                            if (headerMap['FarmerShare']) newRow.FarmerShare = row[headerMap['FarmerShare']];
                            if (headerMap['GovernmentShare']) newRow.GovernmentShare = row[headerMap['GovernmentShare']];
                            if (headerMap['Planting']) newRow.Planting = row[headerMap['Planting']];

                            // Calculate AmountCover based on InsuranceLine
                            const insLine = String(newRow.InsuranceLine || '').toLowerCase();

                            // CropType Logic: Override with AnimalType if Livestock or BoatType if Banca
                            if (insLine.includes('livestock')) {
                                // Use AnimalType if available
                                const animalType = row[headerMap['AnimalType']] || row['AnimalType'];
                                if (animalType) {
                                    newRow.CropType = animalType;
                                }
                            } else if (insLine.includes('banca')) {
                                // Use BoatType if available
                                const boatType = row[headerMap['BoatType']] || row['BoatType'];
                                if (boatType) {
                                    newRow.CropType = boatType;
                                }
                            }



                            // Class Logic: Combine BoatMaterial and Classification
                            const boatMaterial = row[headerMap['BoatMaterial']] || row['BoatMaterial'] || '';
                            const classification = row[headerMap['Classification']] || row['Classification'] || '';

                            // If both exist, join with ' / '. If only one exists, use that.
                            if (boatMaterial && classification) {
                                newRow.Class = `${boatMaterial} / ${classification}`;
                            } else {
                                newRow.Class = boatMaterial || classification || '';
                            }

                            // Amount Cover Logic - Reverted to strict checks as requested
                            if (insLine.includes('crop')) {
                                newRow.AmountCover = row['Crop AmountCover'] || row['AmountCover'] || '';
                            } else if (insLine.includes('livestock')) {
                                newRow.AmountCover = row['Livestock AmountCover'] || row['AmountCover'] || '';
                            } else if (insLine.includes('adss') || insLine.includes('fisheries')) {
                                newRow.AmountCover = row['ADSS AmountCover'] || row['AmountCover'] || '';
                            } else if (insLine.includes('banca')) {
                                newRow.AmountCover = row['Banca AmountCover'] || row['AmountCover'] || '';
                            } else {
                                // Fallback to generic AmountCover if available
                                newRow.AmountCover = row['AmountCover'] || '';
                            }

                            return newRow;
                        });

                        newRecords = [...newRecords, ...transformedData];
                        resolve();
                    },
                    error: function (err) {
                        console.error(`Error parsing ${file.name}:`, err);
                        resolve(); // Continue even if one fails
                    }
                });
            });
        }

        // Append new records to existing data
        csvData = [...csvData, ...newRecords];

        // Save to IndexedDB
        await saveToStorage();

        // Restore Dropzone UI
        dropZone.innerHTML = `
            <div class="upload-content">
                <div class="upload-icon">📁</div>
                <h3>Upload CSV File(s)</h3>
                <p>Drag & drop your files here or click to browse</p>
                <input type="file" id="csv-input" accept=".csv" multiple hidden>
                <button class="btn btn-primary" onclick="document.getElementById('csv-input').click()">Select File(s)</button>
            </div>
        `;



        initializeDashboard();
    }

    async function saveToStorage() {
        try {
            await idbKeyval.set('agriData_csv', csvData);
            await idbKeyval.set('agriData_headers', headers);
        } catch (err) {
            console.error('Save failed:', err);
            alert('Failed to save data to storage. Quota might be exceeded.');
        }
    }

    async function loadFromStorage() {
        try {
            const storedData = await idbKeyval.get('agriData_csv');
            const storedHeaders = await idbKeyval.get('agriData_headers');

            if (storedData && storedData.length > 0) {
                csvData = storedData.map(row => {
                    if (row['Farmer Name'] === undefined) {
                        const lName = row.LastName || '';
                        const fName = row.FirstName || '';
                        const mName = row.MiddlName || '';
                        const extName = row.ExtName || '';

                        let fullName = `${lName}`;
                        if (fullName && (fName || mName || extName)) fullName += ', ';
                        let givenNames = [fName, mName, extName].filter(n => n.trim() !== '').join(' ');
                        fullName += givenNames;

                        row['Farmer Name'] = fullName.trim();
                    }
                    return row;
                });
                // Ensure Farmer Name is in headers if not present
                headers = storedHeaders || Object.keys(storedData[0]);
                if (!headers.includes('Farmer Name')) headers.push('Farmer Name');
                initializeDashboard();
            }
        } catch (err) {
            console.error('Load failed:', err);
        }
    }

    async function clearStorage() {
        await idbKeyval.clear();
        csvData = [];
        headers = [];
        selectedRowIds = [];
        dashboard.classList.add('hidden');
        dropZone.classList.remove('hidden');
        rowCountBadge.textContent = '0 Rows';
    }

    function initializeDashboard() {
        dashboard.classList.remove('hidden');
        dropZone.classList.add('hidden');

        // Generate Summary
        updateSummary();

        // Initial Filter & Render
        updateDashboard();
    }

    function updateSummary() {
        const totalFarmers = new Set(csvData.map(d => String(d.FarmersID || '').trim()).filter(Boolean)).size;

        // Count unique FarmIDs for Total Farms. Do not count blank FarmIDs.
        const totalFarms = new Set(csvData.map(d => String(d.FarmID || '').trim()).filter(Boolean)).size;

        const sumField = (field) => csvData.reduce((acc, curr) => {
            const val = parseFloat(curr[field]);
            return acc + (isNaN(val) ? 0 : val);
        }, 0);

        const totalArea = sumField('Area');


        const totalLivestock = csvData.filter(d => String(d.InsuranceLine).toLowerCase().includes('livestock')).length;
        const totalBoats = csvData.filter(d => String(d.InsuranceLine).toLowerCase().includes('banca')).length;
        const totalCrops = csvData.filter(d => String(d.InsuranceLine).toLowerCase().includes('crop')).length;
        const totalADSS = csvData.filter(d => String(d.InsuranceLine).toUpperCase().includes('ADSS')).length;

        const metrics = [
            { label: 'Unique Farmers', value: totalFarmers, icon: '👨‍🌾' },
            { label: 'Total Area (Ha)', value: totalArea.toFixed(2), icon: '🗺️' },
            { label: 'Total Crops', value: totalCrops, icon: '🌾' },
            { label: 'Total Livestock', value: totalLivestock, icon: '🐄' },
            { label: 'Total Boats', value: totalBoats, icon: '🚤' },
            { label: 'Total ADSS', value: totalADSS, icon: '📡' }
        ];

        summaryCards.innerHTML = metrics.map(m => `
            <div class="card">
                <div class="card-title">${m.icon} ${m.label}</div>
                <div class="card-value">${m.value}</div>
            </div>
        `).join('');
    }

    function renderTable(data, currentInsuranceFilter = '') {
        if (!headers.length) return;

        // Define exact columns to show.
        // If the 'All Insurance Lines' filter is set to 'rice', 'corn', or 'crop' (HVC Crop), show 'Planting' instead of 'Class'
        const showPlantingColumn = currentInsuranceFilter === 'rice' || currentInsuranceFilter === 'corn' || currentInsuranceFilter === 'crop';

        const targetColumns = [
            'Checkbox', 'No.', 'FarmersID', 'RSBSAID', 'Farmer Name',
            'MunFarmer', 'BrgyFarmer', 'StFarmer', 'InsuranceLine',
            ...(showPlantingColumn ? ['Planting'] : []),
            'CropType',
            ...(!showPlantingColumn ? ['Class'] : []),
            'AmountCover'
        ];

        // Filter headers to only include available target columns
        // This handles case-insensitive matching to find the actual header name in the CSV
        // For 'Checkbox', we just return it as is.
        const displayHeaders = targetColumns.map(target => {
            if (target === 'Checkbox') return 'Checkbox';
            return headers.find(h => h.toLowerCase().trim() === target.toLowerCase().trim()) || target;
        });

        tableHead.innerHTML = `<tr>${displayHeaders.map(h => {
            if (h === 'Checkbox') {
                return `<th style="width: 40px; text-align: center;"><input type="checkbox" id="select-all"></th>`;
            }
            const extraStyle = h === 'Farmer Name' ? 'style="min-width: 250px; width: 250px;"' : '';
            return `<th title="${h}" ${extraStyle}>${h}</th>`;
        }).join('')}</tr>`;

        if (data.length === 0) {
            tableBody.innerHTML = `<tr><td colspan="${displayHeaders.length}" style="text-align: center; color: var(--text-muted); font-style: italic; padding: 2rem; font-weight: bold; font-size: 1.1rem;">No Commodity Insured</td></tr>`;
            return;
        }

        const displayData = data.slice(0, 500);
        tableBody.innerHTML = displayData.map(row => `
            <tr>
                ${displayHeaders.map(h => {
            if (h === 'Checkbox') {
                return `<td style="text-align: center;"><input type="checkbox" class="row-checkbox" data-row-id="${row['No.']}" ${selectedRowIds.includes(String(row['No.'])) ? 'checked' : ''}></td>`;
            }

            let content = row[h] === null || row[h] === undefined ? '' : row[h];

            // Format AmountCover with commas
            if (h === 'AmountCover' && content !== '') {
                const num = parseFloat(String(content).replace(/,/g, ''));
                if (!isNaN(num)) {
                    content = num.toLocaleString('en-US');
                }
            }

            const extraStyle = h === 'Farmer Name' ? 'style="white-space: nowrap; min-width: 250px; width: 250px;"' : '';
            return `<td title="${content}" ${extraStyle}>${content}</td>`;
        }).join('')}
            </tr>
        `).join('');

        if (data.length > 500) {
            const infoRow = document.createElement('tr');
            infoRow.innerHTML = `<td colspan="${displayHeaders.length}" style="text-align: center; color: var(--text-muted); font-style: italic; padding: 2rem;">Showing first 500 rows. Use search to find specific records.</td>`;
            tableBody.appendChild(infoRow);
        }
    }

    // --- Dynamic Filter Logic ---

    function updateDashboard() {
        const searchTerm = tableSearch.value.toLowerCase();
        const insuranceValue = insuranceFilter.value.toLowerCase();
        const classValue = classFilter.value.toLowerCase(); // Added Class Filter value
        const cropValue = cropFilter.value.toLowerCase();
        const dateFrom = dateFromFilter.value;
        const dateTo = dateToFilter.value;

        // 1. Calculate Table Data (Intersection of ALL filters)
        const tableData = csvData.filter(row => {
            return matchesSearch(row, searchTerm) &&
                matchesInsurance(row, insuranceValue) &&
                matchesClass(row, classValue) && // Added Class Filter to tableData
                matchesCrop(row, cropValue) &&
                matchesDate(row, dateFrom, dateTo);
        });

        // 2. Update Insurance Options (Based on Search + Class + Crop, IGNORING Insurance)
        const insuranceAvailableData = csvData.filter(row => {
            return matchesSearch(row, searchTerm) &&
                matchesClass(row, classValue) && // Added Class Filter
                matchesCrop(row, cropValue) &&
                matchesDate(row, dateFrom, dateTo);
        });
        populateFilterOptions(insuranceFilter, 'InsuranceLine', 'All Insurance Lines', insuranceAvailableData);

        // 3. Update Class Options (Based on Search + Insurance + Crop, IGNORING Class)
        const classAvailableData = csvData.filter(row => {
            return matchesSearch(row, searchTerm) &&
                matchesInsurance(row, insuranceValue) &&
                matchesCrop(row, cropValue) &&
                matchesDate(row, dateFrom, dateTo);
        });
        populateFilterOptions(classFilter, 'Class', 'All Classes', classAvailableData); // Added Class Filter population

        // 4. Update Crop Options (Based on Search + Insurance + Class, IGNORING Crop)
        const cropAvailableData = csvData.filter(row => {
            return matchesSearch(row, searchTerm) &&
                matchesInsurance(row, insuranceValue) &&
                matchesClass(row, classValue) &&
                matchesDate(row, dateFrom, dateTo);
        });
        populateFilterOptions(cropFilter, 'CropType', 'All Crop Types', cropAvailableData);

        // Render Table
        renderTable(tableData, insuranceValue);

        // Update bulk actions bar based on selected rows
        if (typeof updateBulkBar === 'function') {
            updateBulkBar();
        }

        rowCountBadge.textContent = `${tableData.length} Match(es)`;

        // Check if ANY filter is active
        const isSearchActive = searchTerm !== '';
        const isInsuranceActive = insuranceValue !== '' && insuranceValue !== 'all';
        const isClassActive = classValue !== '' && classValue !== 'all';
        const isCropActive = cropValue !== '' && cropValue !== 'all';
        const isDateActive = dateFrom !== '' || dateTo !== '';

        const noFiltersActive = !isSearchActive && !isInsuranceActive && !isClassActive && !isCropActive && !isDateActive;

        // Auto-select report
        if (noFiltersActive) {
            // If no filters are applied, prompt user to filter first
            reportTypeSelector.value = 'none';
        } else if (tableData.length > 0) {
            // If filters are applied, select based on majority InsuranceLine or CropType
            const counts = {};
            let maxCount = 0;
            let majorityLine = '';

            const cropCounts = {};
            let maxCropCount = 0;
            let majorityCrop = '';

            tableData.forEach(row => {
                // Count InsuranceLine
                const line = String(row.InsuranceLine || '').toLowerCase();
                if (line) {
                    counts[line] = (counts[line] || 0) + 1;
                    if (counts[line] > maxCount) {
                        maxCount = counts[line];
                        majorityLine = line;
                    }
                }

                // Count CropType
                const crop = String(row.CropType || '').toLowerCase();
                if (crop) {
                    cropCounts[crop] = (cropCounts[crop] || 0) + 1;
                    if (cropCounts[crop] > maxCropCount) {
                        maxCropCount = cropCounts[crop];
                        majorityCrop = crop;
                    }
                }
            });

            // Auto-selection prioritization
            const isRiceCornCrop = majorityCrop.includes('rice') || majorityCrop.includes('palay') || majorityCrop.includes('corn');
            const isRiceCornLine = majorityLine.includes('rice') || majorityLine.includes('corn');
            const isHvcFilterSelected = insuranceFilter.value.toLowerCase() === 'crop';

            const hasNoSpecializedMajorityLine = !majorityLine.includes('livestock') && !majorityLine.includes('banca') && !majorityLine.includes('noncrop') && !majorityLine.includes('adss') && !majorityLine.includes('tir');
            const hasNoRiceCornMajorityCrop = !majorityCrop.includes('rice') && !majorityCrop.includes('palay') && !majorityCrop.includes('corn');

            // 1. If explicitly HVC filter, OR if majority is neither specialized line nor rice/corn
            if (isHvcFilterSelected || (hasNoSpecializedMajorityLine && hasNoRiceCornMajorityCrop && majorityLine !== '')) {
                reportTypeSelector.value = 'hvc';
            }
            // 2. If explicit Rice/Corn Line, OR if clearly a Rice/Corn dataset based on CropType majority 
            else if (isRiceCornLine || (isRiceCornCrop && maxCropCount >= maxCount * 0.5)) {
                reportTypeSelector.value = 'rice-corn';
            }
            // 3. Standard Fallbacks based on InsuranceLine
            else if (majorityLine.includes('livestock')) {
                reportTypeSelector.value = 'livestock';
            } else if (majorityLine.includes('adss') || majorityLine.includes('tir')) {
                reportTypeSelector.value = 'tir';
            } else if (majorityLine.includes('banca') || majorityLine.includes('noncrop')) {
                reportTypeSelector.value = 'noncrop';
            }
        }
    }

    // Helper match functions to avoid repetition
    const matchesSearch = (row, term) => Object.values(row).some(val => String(val).toLowerCase().includes(term));
    const matchesInsurance = (row, val) => {
        if (val === '' || val === 'all') return true;

        // Special mapping: if user selected "corn" or "rice" in Insurance filter, 
        // we actually check the CropType column instead.
        if (val === 'corn') return String(row.CropType || '').toLowerCase() === 'corn';
        if (val === 'rice') {
            const cropTypeStr = String(row.CropType || '').toLowerCase();
            return cropTypeStr === 'rice' || cropTypeStr === 'palay';
        }

        // Special mapping: if user selected "crop" (HVC Crop), ensure we EXCLUDE rice and corn
        if (val === 'crop') {
            const isCropLine = String(row.InsuranceLine || '').toLowerCase() === 'crop';
            if (!isCropLine) return false;

            const cropTypeStr = String(row.CropType || '').toLowerCase();
            const isRiceOrCorn = cropTypeStr === 'rice' || cropTypeStr === 'palay' || cropTypeStr === 'corn';
            return !isRiceOrCorn;
        }

        // Standard behavior: check InsuranceLine column
        return String(row.InsuranceLine || '').toLowerCase() === val;
    };
    const matchesClass = (row, val) => val === '' || val === 'all' || String(row.Class || '').toLowerCase() === val;
    const matchesCrop = (row, val) => val === '' || val === 'all' || String(row.CropType || '').toLowerCase() === val;
    const matchesDate = (row, fromDateStr, toDateStr) => {
        if (!fromDateStr && !toDateStr) return true; // No date filter applied

        const rawPlantingDate = row.Planting;
        if (!rawPlantingDate) return false; // Date required but not present

        const plantingDate = new Date(rawPlantingDate);
        if (isNaN(plantingDate.getTime())) return false; // Invalid date format

        // Normalize time to 00:00:00 for accurate comparison if we only care about the date part.
        plantingDate.setHours(0, 0, 0, 0);

        if (fromDateStr) {
            const fromDate = new Date(fromDateStr);
            fromDate.setHours(0, 0, 0, 0);
            if (plantingDate < fromDate) return false;
        }

        if (toDateStr) {
            const toDate = new Date(toDateStr);
            toDate.setHours(0, 0, 0, 0);
            if (plantingDate > toDate) return false;
        }

        return true;
    };


    // Helper to populate select elements with dynamic data
    function populateFilterOptions(selectElement, fieldName, defaultMetadata, dataSubset) {
        // Get unique values from the SUBSET of data
        const values = new Set(dataSubset.map(d => String(d[fieldName] || '').trim()).filter(l => l));

        // Explicitly inject "Rice" and "Corn" as options into the Insurance Line dropdown
        if (fieldName === 'InsuranceLine') {
            values.add('Rice');
            values.add('Corn');
        }

        let sortedValues;

        if (fieldName === 'InsuranceLine') {
            // Apply custom sort order for Insurance Line
            const customOrder = ['Rice', 'Corn', 'ADSS', 'Livestock', 'Crop', 'Banca'];
            sortedValues = Array.from(values).sort((a, b) => {
                const indexA = customOrder.indexOf(a);
                const indexB = customOrder.indexOf(b);
                // If both are in custom list, sort by index
                if (indexA !== -1 && indexB !== -1) return indexA - indexB;
                // If only one is in custom list, prioritize it
                if (indexA !== -1) return -1;
                if (indexB !== -1) return 1;
                // If neither in custom list, sort alphabetically
                return a.localeCompare(b);
            });
        } else {
            // Default alphabetical sort for other fields
            sortedValues = Array.from(values).sort();
        }

        // Save current selection (to try and preserve it)
        const currentSelection = selectElement.value;

        // Reset options
        selectElement.innerHTML = `<option value="">${defaultMetadata}</option>`;

        sortedValues.forEach(val => {
            const option = document.createElement('option');
            option.value = val;

            // Text replacement for 'Crop' to display as 'HVC Crop'
            if (fieldName === 'InsuranceLine' && val === 'Crop') {
                option.textContent = 'HVC Crop';
            } else {
                option.textContent = val;
            }

            selectElement.appendChild(option);
        });

        // Restore selection if it works with the new data
        // If the current selection is NO LONGER valid (not in the filtered list),
        // it will naturally fall back to the first option ("") which is correct behavior for dependent filters.
        if (sortedValues.includes(currentSelection)) {
            selectElement.value = currentSelection;
        }
    }

    // Event Listeners - All point to the central update function
    tableSearch.addEventListener('input', updateDashboard);
    insuranceFilter.addEventListener('change', updateDashboard);
    classFilter.addEventListener('change', updateDashboard);
    cropFilter.addEventListener('change', updateDashboard);
    dateFromFilter.addEventListener('change', updateDashboard);
    dateToFilter.addEventListener('change', updateDashboard);

    // Checkbox Event Delegation for Bulk Actions
    const bulkActionsBar = document.getElementById('bulk-actions-bar');
    const selectedCount = document.getElementById('selected-count');

    function updateBulkBar() {
        if (selectedRowIds.length > 0) {
            bulkActionsBar.classList.remove('hidden');
            selectedCount.textContent = `${selectedRowIds.length} Row${selectedRowIds.length > 1 ? 's' : ''} Selected`;
        } else {
            bulkActionsBar.classList.add('hidden');
        }
    }

    tableHead.addEventListener('change', (e) => {
        if (e.target.id === 'select-all') {
            const isChecked = e.target.checked;
            const rowCheckboxes = document.querySelectorAll('.row-checkbox');
            rowCheckboxes.forEach(cb => {
                cb.checked = isChecked;
                const rowId = String(cb.dataset.rowId);
                if (isChecked) {
                    if (!selectedRowIds.includes(rowId)) selectedRowIds.push(rowId);
                } else {
                    selectedRowIds = selectedRowIds.filter(id => id !== rowId);
                }
            });
            updateBulkBar();
        }
    });

    tableBody.addEventListener('change', (e) => {
        if (e.target.classList.contains('row-checkbox')) {
            const rowId = String(e.target.dataset.rowId);
            if (e.target.checked) {
                if (!selectedRowIds.includes(rowId)) selectedRowIds.push(rowId);
            } else {
                selectedRowIds = selectedRowIds.filter(id => id !== rowId);
            }

            const rowCheckboxes = document.querySelectorAll('.row-checkbox');
            const allChecked = rowCheckboxes.length > 0 && Array.from(rowCheckboxes).every(cb => cb.checked);
            const selectAllCb = document.getElementById('select-all');
            if (selectAllCb) selectAllCb.checked = allChecked;
            updateBulkBar();
        }
    });

    // Reset Filters Button
    if (resetFiltersBtn) {
        resetFiltersBtn.addEventListener('click', () => {
            selectedRowIds = [];
            const selectAllCb = document.getElementById('select-all');
            if (selectAllCb) selectAllCb.checked = false;
            updateBulkBar();

            tableSearch.value = '';
            insuranceFilter.value = 'all';
            classFilter.value = 'all';
            cropFilter.value = 'all';
            dateFromFilter.value = '';
            dateToFilter.value = '';
            updateDashboard();
        });
    }

    // Print Report Button
    if (printBtn) {
        printBtn.addEventListener('click', () => generatePCICReport());
    }

    const printSelectedBtn = document.getElementById('print-selected-btn');
    if (printSelectedBtn) {
        printSelectedBtn.addEventListener('click', () => {
            generatePCICReport(selectedRowIds);
        });
    }

    const deselectAllBtn = document.getElementById('deselect-all-btn');
    if (deselectAllBtn) {
        deselectAllBtn.addEventListener('click', () => {
            selectedRowIds = [];
            document.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = false);
            const selectAllCb = document.getElementById('select-all');
            if (selectAllCb) selectAllCb.checked = false;
            updateBulkBar();
        });
    }

    // --- MODAL & PREVIEW LOGIC ---
    const modalPrintBtn = document.getElementById('modal-print-btn');
    const modalCloseBtn = document.getElementById('modal-close-btn');
    const printModal = document.getElementById('print-modal');

    if (modalPrintBtn) {
        modalPrintBtn.addEventListener('click', () => {
            window.print();
        });
    }

    if (modalCloseBtn) {
        modalCloseBtn.addEventListener('click', () => {
            if (printModal) printModal.classList.add('hidden');
        });
    }

    if (printModal) {
        printModal.addEventListener('click', (e) => {
            if (e.target === printModal) {
                printModal.classList.add('hidden');
            }
        });
    }

    function generatePCICReport(selectedIds = null) {
        const searchTerm = tableSearch.value.toLowerCase();
        const insuranceValue = insuranceFilter.value.toLowerCase();
        const classValue = classFilter.value.toLowerCase();
        const cropValue = cropFilter.value.toLowerCase();
        const dateFrom = dateFromFilter.value;
        const dateTo = dateToFilter.value;

        let reportData = csvData.filter(row => {
            return matchesSearch(row, searchTerm) &&
                matchesInsurance(row, insuranceValue) &&
                matchesClass(row, classValue) &&
                matchesCrop(row, cropValue) &&
                matchesDate(row, dateFrom, dateTo);
        });

        if (selectedIds && selectedIds.length > 0) {
            reportData = reportData.filter(row => selectedIds.includes(String(row['No.'])));
        }

        if (reportData.length === 0) {
            alert('No data to print!');
            return;
        }

        // Calculate Totals using logic similar to updateSummary but on filtered data
        const totalFarmers = new Set(reportData.map(d => String(d.FarmersID || '').trim()).filter(Boolean)).size;

        // Count unique FarmIDs for Total Farms. Do not count blank FarmIDs.
        const farmIds = reportData.map(d => String(d.FarmID || '').trim()).filter(Boolean);
        const totalFarms = new Set(farmIds).size;

        const totalArea = reportData.reduce((acc, curr) => acc + (parseFloat(curr.Area) || 0), 0);

        // Financials (Summing mapped fields)
        const farmerShare = reportData.reduce((acc, curr) => {
            const val = String(curr.FarmerShare || 0).replace(/,/g, '');
            return acc + (parseFloat(val) || 0);
        }, 0);

        const govShare = reportData.reduce((acc, curr) => {
            const val = String(curr.GovernmentShare || 0).replace(/,/g, '');
            return acc + (parseFloat(val) || 0);
        }, 0);

        const totalAmountCover = reportData.reduce((acc, curr) => {
            const val = String(curr.AmountCover || 0).replace(/,/g, '');
            return acc + (parseFloat(val) || 0);
        }, 0);

        const totalGrossPremium = totalAmountCover * 0.10;

        // Populate Template (Values for Inputs)

        // Helper to set value safely to inputs or innerText
        const setVal = (id, val) => {
            const el = document.getElementById(id);
            if (el) {
                if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT') {
                    el.value = val;
                } else {
                    el.textContent = val;
                }
            }
        };

        // Handle Report Type Visibility
        const reportType = reportTypeSelector ? reportTypeSelector.value : 'rice-corn';
        const isTIR = reportType === 'tir';
        const isHVC = reportType === 'hvc';
        const isFisheries = reportType === 'fisheries';
        const isLivestock = reportType === 'livestock';
        const isNonCrop = reportType === 'noncrop';
        const ricencornTemplate = document.getElementById('ricencorn-preprocessing-slip');
        const tirTemplate = document.getElementById('tir-report-template');
        const adssTemplate = document.getElementById('adss-slip-template');
        const hvcTemplate = document.getElementById('hvc-preprocessing-slip');
        const fisheriesTemplate = document.getElementById('fisheries-preprocessing-slip');
        const livestockTemplate = document.getElementById('livestock-preprocessing-slip');
        const noncropTemplate = document.getElementById('noncrop-preprocessing-slip');

        // Hide all templates first
        if (ricencornTemplate) ricencornTemplate.classList.add('hidden');
        if (tirTemplate) tirTemplate.classList.add('hidden');
        if (adssTemplate) adssTemplate.classList.add('hidden');
        if (hvcTemplate) hvcTemplate.classList.add('hidden');
        if (fisheriesTemplate) fisheriesTemplate.classList.add('hidden');
        if (livestockTemplate) livestockTemplate.classList.add('hidden');
        if (noncropTemplate) noncropTemplate.classList.add('hidden');

        if (isTIR) {
            if (tirTemplate) tirTemplate.classList.remove('hidden');
            if (adssTemplate) adssTemplate.classList.remove('hidden');

            renderTIRReport(reportData);
            renderADSSReport(reportData);

            if (printModal) printModal.classList.remove('hidden');

        } else if (isHVC) {
            if (hvcTemplate) hvcTemplate.classList.remove('hidden');

            // Auto-populate Farmers Group from CSV
            const hvcFarmersGroup = document.getElementById('hvc-farmers-group');
            if (hvcFarmersGroup && reportData.length > 0) {
                const firstRow = reportData[0];
                hvcFarmersGroup.textContent = firstRow['FarmersGroup'] || firstRow['GroupName'] || '';
            }

            // Auto-populate Date Received
            const hvcDateReceived = document.getElementById('hvc-date-received');
            if (hvcDateReceived) {
                const today = new Date();
                hvcDateReceived.textContent = today.toLocaleDateString('en-US'); // Will show MM/DD/YYYY
            }

            // Auto-populate Crop/Variety
            const hvcCropVariety = document.getElementById('hvc-crop-variety');
            if (hvcCropVariety && reportData.length > 0) {
                const uniqueCropTypes = [...new Set(reportData.map(row => String(row['CropType'] || '')).filter(val => val.trim() !== ''))];
                hvcCropVariety.textContent = uniqueCropTypes.join(' & ');
            }

            // Calculate and populate As Submitted totals
            const uniqueFarmersHvc = new Set(reportData.map(row => String(row['FarmersID'] || '')).filter(val => val.trim() !== ''));
            const totalFarmsHvc = new Set(reportData.map(row => String(row['FarmID'] || '').trim()).filter(Boolean)).size;
            const totalAreaHvc = reportData.reduce((sum, row) => sum + (parseFloat(row['Area']) || 0), 0);

            setVal('hvc-as-sub-farmers', uniqueFarmersHvc.size);
            setVal('hvc-as-sub-farms', totalFarmsHvc);

            // Format Area as a locale string with 2 or 4 decimals based on values, but typically 2 is standard
            const formattedArea = totalAreaHvc.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 4 });
            setVal('hvc-as-sub-area', formattedArea);

            // Total Sum Insured (based on 'Crop AmountCover')
            const totalSumInsuredHvc = reportData.reduce((sum, row) => {
                const rawAmount = row['Crop AmountCover'];
                // Clean comma formatted strings if any, then parse
                const amount = parseFloat(String(rawAmount || 0).replace(/,/g, ''));
                return sum + (isNaN(amount) ? 0 : amount);
            }, 0);
            setVal('hvc-total-sum-insured', totalSumInsuredHvc.toLocaleString('en-US', { minimumFractionDigits: 2 }));


            // Activity Dates
            const hvcDateReceivedText = new Date().toLocaleDateString('en-US');
            setVal('hvc-act1-date', hvcDateReceivedText);
            setVal('hvc-act2-date', hvcDateReceivedText);

            if (printModal) printModal.classList.remove('hidden');

        } else if (isFisheries) {
            if (fisheriesTemplate) fisheriesTemplate.classList.remove('hidden');

            // Auto-populate Farmers Group from CSV for Fisheries
            const fisheriesFarmersGroup = document.getElementById('fisheries-farmers-group');
            if (fisheriesFarmersGroup && reportData.length > 0) {
                const firstRow = reportData[0];
                fisheriesFarmersGroup.textContent = firstRow['FarmersGroup'] || firstRow['GroupName'] || '';
            }

            // Activity Dates for Fisheries
            const fishDateReceivedText = new Date().toLocaleDateString('en-US');
            setVal('fish-act1-date', fishDateReceivedText);
            setVal('fish-act2-date', fishDateReceivedText);

            if (printModal) printModal.classList.remove('hidden');

        } else if (isLivestock) {
            if (livestockTemplate) livestockTemplate.classList.remove('hidden');

            // Auto-populate Farmers Group from CSV for Livestock
            const livestockFarmersGroup = document.getElementById('livestock-farmers-group');
            if (livestockFarmersGroup && reportData.length > 0) {
                const firstRow = reportData[0];
                livestockFarmersGroup.textContent = firstRow['FarmersGroup'] || firstRow['GroupName'] || '';
            }

            // Total Sum Insured (based on 'Livestock AmountCover')
            const totalSumInsuredLivestock = reportData.reduce((sum, row) => {
                const rawAmount = row['Livestock AmountCover'];
                // Clean comma formatted strings if any, then parse
                const amount = parseFloat(String(rawAmount || 0).replace(/,/g, ''));
                return sum + (isNaN(amount) ? 0 : amount);
            }, 0);
            setVal('livestock-total-sum-insured', totalSumInsuredLivestock.toLocaleString('en-US', { minimumFractionDigits: 2 }));

            // Number of Heads (total rows shown in the slip)
            setVal('livestock-number-of-heads', reportData.length);

            // Number of Unique Farmers
            const uniqueFarmersLivestock = new Set(reportData.map(row => String(row['FarmersID'] || '')).filter(val => val.trim() !== ''));
            setVal('livestock-number-of-farmers', uniqueFarmersLivestock.size);

            // Animals/ Species (Majority CropType + Majority Class)
            const livestockCropCounts = {};
            let livestockMaxCropCount = 0;
            let livestockMajorityCrop = '';

            const livestockClassCounts = {};
            let livestockMaxClassCount = 0;
            let livestockMajorityClass = '';

            reportData.forEach(row => {
                const crop = String(row.CropType || '').trim();
                const cls = String(row.Class || '').trim();

                if (crop) {
                    livestockCropCounts[crop] = (livestockCropCounts[crop] || 0) + 1;
                    if (livestockCropCounts[crop] > livestockMaxCropCount) {
                        livestockMaxCropCount = livestockCropCounts[crop];
                        livestockMajorityCrop = crop;
                    }
                }

                if (cls) {
                    livestockClassCounts[cls] = (livestockClassCounts[cls] || 0) + 1;
                    if (livestockClassCounts[cls] > livestockMaxClassCount) {
                        livestockMaxClassCount = livestockClassCounts[cls];
                        livestockMajorityClass = cls;
                    }
                }
            });

            const speciesDisplay = livestockMajorityCrop && livestockMajorityClass ?
                `${livestockMajorityCrop} (${livestockMajorityClass})` :
                (livestockMajorityCrop || livestockMajorityClass || '');

            const speciesElement = document.getElementById('livestock-animal-species');
            if (speciesElement) speciesElement.textContent = speciesDisplay;

            // Auto-populate Date Received
            const livestockDateReceived = document.getElementById('livestock-date-received');
            const currentFormattedDate = new Date().toLocaleDateString('en-US');
            if (livestockDateReceived) {
                livestockDateReceived.textContent = currentFormattedDate;
            }

            // Activity Dates
            setVal('livestock-act1-date', currentFormattedDate);
            setVal('livestock-act2-date', currentFormattedDate);

            if (printModal) printModal.classList.remove('hidden');

        } else if (isNonCrop) {
            if (noncropTemplate) noncropTemplate.classList.remove('hidden');

            // Auto-populate Farmers Group from CSV for Non-Crop
            const noncropFarmersGroup = document.getElementById('noncrop-farmers-group');
            if (noncropFarmersGroup && reportData.length > 0) {
                const firstRow = reportData[0];
                noncropFarmersGroup.textContent = firstRow['FarmersGroup'] || firstRow['GroupName'] || '';
            }

            // Total Sum Insured (based on 'Banca AmountCover')
            const totalSumInsuredNonCrop = reportData.reduce((sum, row) => {
                const rawAmount = row['Banca AmountCover'];
                // Clean comma formatted strings if any, then parse
                const amount = parseFloat(String(rawAmount || 0).replace(/,/g, ''));
                return sum + (isNaN(amount) ? 0 : amount);
            }, 0);
            setVal('noncrop-total-sum-insured', totalSumInsuredNonCrop.toLocaleString('en-US', { minimumFractionDigits: 2 }));

            // Number of Unique Farmers
            const uniqueFarmersNonCrop = new Set(reportData.map(row => String(row['FarmersID'] || '')).filter(val => val.trim() !== ''));
            setVal('noncrop-number-of-farmers', uniqueFarmersNonCrop.size);

            // Number of Units (total rows shown in the slip)
            setVal('noncrop-number-of-units', reportData.length);

            // NCAAI Unit (Majority CropType + Majority Class)
            const noncropCropCounts = {};
            let noncropMaxCropCount = 0;
            let noncropMajorityCrop = '';

            const noncropClassCounts = {};
            let noncropMaxClassCount = 0;
            let noncropMajorityClass = '';

            reportData.forEach(row => {
                const crop = String(row.CropType || '').trim();
                const cls = String(row.Class || '').trim();

                if (crop) {
                    noncropCropCounts[crop] = (noncropCropCounts[crop] || 0) + 1;
                    if (noncropCropCounts[crop] > noncropMaxCropCount) {
                        noncropMaxCropCount = noncropCropCounts[crop];
                        noncropMajorityCrop = crop;
                    }
                }

                if (cls) {
                    noncropClassCounts[cls] = (noncropClassCounts[cls] || 0) + 1;
                    if (noncropClassCounts[cls] > noncropMaxClassCount) {
                        noncropMaxClassCount = noncropClassCounts[cls];
                        noncropMajorityClass = cls;
                    }
                }
            });

            const noncropSpeciesDisplay = noncropMajorityCrop && noncropMajorityClass ?
                `${noncropMajorityCrop} (${noncropMajorityClass})` :
                (noncropMajorityCrop || noncropMajorityClass || '');

            const noncropSpeciesElement = document.getElementById('noncrop-ncaai-unit');
            if (noncropSpeciesElement) noncropSpeciesElement.textContent = noncropSpeciesDisplay;

            // Auto-populate Date Received
            const noncropDateReceived = document.getElementById('noncrop-date-received');
            const currentFormattedDate = new Date().toLocaleDateString('en-US');
            if (noncropDateReceived) {
                noncropDateReceived.textContent = currentFormattedDate; // Will show MM/DD/YYYY
            }

            // Activity Dates
            setVal('noncrop-act1-date', currentFormattedDate);
            setVal('noncrop-act2-date', currentFormattedDate);

            if (printModal) printModal.classList.remove('hidden');

        } else {
            if (ricencornTemplate) ricencornTemplate.classList.remove('hidden');

            // Populate Rice/Corn Template
            setVal('print-date-received', new Date().toLocaleDateString());

            // Rice/Corn Check Logic (Put '✔' centered on the line)
            const isRice = insuranceValue === 'rice' || insuranceValue.includes('rice');
            const isCorn = insuranceValue === 'corn' || insuranceValue.includes('corn');

            const riceLine = document.getElementById('print-rice-line');
            const cornLine = document.getElementById('print-corn-line');

            // Reset check logic and ensure visual center alignment
            if (riceLine) {
                riceLine.value = '';
                riceLine.style.textAlign = 'center';
            }
            if (cornLine) {
                cornLine.value = '';
                cornLine.style.textAlign = 'center';
            }

            if (riceLine && isRice) riceLine.value = '✔';
            if (cornLine && isCorn) cornLine.value = '✔';

            // Clear/Reset Placeholders
            const fieldsToClear = [
                'print-logbook-no', 'print-phase', 'print-lender-name', 'print-group-name',
                'print-cic-no', 'print-cic-date', 'print-expiry', 'print-li-share',
                'print-or-no', 'print-or-amount', 'print-note',
                'print-farmer-li', 'print-service-fee', 'print-net-remittance',
                'print-over-under', 'print-or-date'
            ];

            fieldsToClear.forEach(id => setVal(id, ''));

            // Map Totals
            setVal('print-total-farmers', totalFarmers);
            setVal('print-total-farms', totalFarms);
            setVal('print-total-area', totalArea.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));

            // Financials
            setVal('print-farmer-share', farmerShare.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
            setVal('print-gov-share', govShare.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
            setVal('print-total-amount-cover', totalAmountCover.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
            setVal('print-gross-premium', totalGrossPremium.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
        }

        // Show Modal
        if (printModal && !isTIR && !isHVC && !isFisheries && !isLivestock && !isNonCrop) {
            printModal.classList.remove('hidden');
        }
    }

    function renderTIRReport(reportData) {
        const tirTableBody = document.getElementById('tir-table-body');
        const tirTotalCover = document.getElementById('tir-total-cover');
        const tirTotalPremium = document.getElementById('tir-total-premium');
        const tirAddress = document.getElementById('tir-address');
        const tirDate = document.getElementById('tir-date');
        const tirGrossIncome = document.getElementById('tir-gross-income');
        const tirNetPremium = document.getElementById('tir-net-premium');

        // Clear existing rows
        if (tirTableBody) {
            tirTableBody.innerHTML = '';
        }

        const currentDate = new Date();
        const nextYearDate = new Date();
        nextYearDate.setFullYear(currentDate.getFullYear() + 1);

        const formatDate = (date) => {
            return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
        };

        const headerDateStr = formatDate(currentDate);
        const expiryDateStr = formatDate(nextYearDate);

        if (tirDate) tirDate.textContent = headerDateStr;

        // Extract address from first row if available (MunFarmer for header)
        if (tirAddress && reportData.length > 0) {
            const row = reportData[0];
            const address = row.MunFarmer || '';
            tirAddress.value = address;
        }

        let totalCover = 0;
        let totalPremium = 0;

        reportData.forEach((row, index) => {
            const name = row['Farmer Name'] || '';
            const address = [row.StFarmer, row.BrgyFarmer].filter(Boolean).join(', ');

            // Period of Cover (Automated: From = Today, To = 1 Year Later)
            const periodFrom = headerDateStr;
            const periodTo = expiryDateStr;

            // Age, Beneficiary, Birthday, Relationship (Placeholder if not in CSV)
            const birthday = row['Birthdate'] || '';
            const age = calculateAge(birthday);
            const beneficiary = row['Beneficiary'] || '';
            const relationship = row['BeneRelationship'] || '';

            let coverStr = String(row.AmountCover || 0).replace(/,/g, '');
            let coverNum = parseFloat(coverStr) || 0;

            let premiumNum = 0;
            if (row.Premium !== undefined && row.Premium !== null && row.Premium !== '') {
                premiumNum = parseFloat(String(row.Premium).replace(/,/g, '')) || 0;
            }

            totalCover += coverNum;
            totalPremium += premiumNum;

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="col-num">${index + 1}</td>
                <td>${name}</td>
                <td>${address}</td>
                <td style="text-align: center;">${birthday}</td>
                <td style="text-align: center;">${age}</td>
                <td>${beneficiary}</td>
                <td style="text-align: center;">${relationship}</td>
                <td style="text-align:right;">${coverNum.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                <td style="text-align: center;">${periodFrom}</td>
                <td style="text-align: center;">${periodTo}</td>
                <td style="text-align:right;">${premiumNum > 0 ? premiumNum.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : ''}</td>
            `;
            tirTableBody.appendChild(tr);
        });

        if (tirTotalCover) tirTotalCover.textContent = totalCover > 0 ? totalCover.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '';
        if (tirTotalPremium) tirTotalPremium.textContent = totalPremium > 0 ? totalPremium.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '';

        // Populate Summary of Premium Remittance
        if (tirGrossIncome) tirGrossIncome.textContent = totalPremium > 0 ? totalPremium.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '';
        if (tirNetPremium) tirNetPremium.textContent = totalPremium > 0 ? totalPremium.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '';
    }

    function renderADSSReport(reportData) {
        if (reportData.length === 0) return;

        // Helper to set value safely
        const setVal = (id, val) => {
            const el = document.getElementById(id);
            if (el) el.value = val;
        };

        const totalFarmers = reportData.length;
        const totalSumInsured = reportData.reduce((acc, curr) => acc + (parseFloat(String(curr.AmountCover || 0).replace(/,/g, '')) || 0), 0);
        const totalPremium = reportData.reduce((acc, curr) => acc + (parseFloat(String(curr.Premium || 0).replace(/,/g, '')) || 0), 0);

        // Dates
        const today = new Date();
        const expiry = new Date();
        expiry.setFullYear(today.getFullYear() + 1);

        const formatDate = (date) => `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;

        // Populate Fields
        setVal('adss-date-received', formatDate(today));
        setVal('adss-agent-name', reportData[0].Lender || ''); // Mapping Lender to Agent as per usual practice
        setVal('adss-group-name', reportData[0].GroupName || '');
        setVal('adss-total-insured', totalFarmers);
        setVal('adss-total-sum', totalSumInsured.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
        setVal('adss-gross-premium', totalPremium.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
        setVal('adss-net-premium', totalPremium.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
        setVal('adss-remitted', totalPremium.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
        setVal('adss-effectivity', formatDate(today));
        setVal('adss-expiry', formatDate(expiry));
    }


});
