document.addEventListener("DOMContentLoaded", () => {
    // UI Elements
    const landingPage = document.getElementById("landingPage");
    const appContainer = document.getElementById("appContainer");
    const excelUpload = document.getElementById("excelUpload");
    const uploadStatus = document.getElementById("uploadStatus");

    const applicantList = document.getElementById("applicantList");
    const searchInput = document.getElementById("searchInput");
    const welcomeMessage = document.getElementById("welcomeMessage");
    const applicantDetails = document.getElementById("applicantDetails");

    let applicantsData = [];
    let originalFilename = 'applicants';

    // Rating state
    let currentApplicantEmail = null;
    let weights = JSON.parse(localStorage.getItem('evalWeights')) || { bsc: 4, msc: 32, research: 16, prof: 8, english: 8, cv: 32 };
    let ratings = JSON.parse(localStorage.getItem('evalRatings')) || {};
    let markedCandidates = JSON.parse(localStorage.getItem('markedCandidates')) || {};
    let reviewerNotes = JSON.parse(localStorage.getItem('reviewerNotes')) || {};

    // Settings Modal & Evaluation UI
    const settingsBtn = document.getElementById("settingsBtn");
    const settingsModal = document.getElementById("settingsModal");
    const closeSettings = document.getElementById("closeSettings");
    const saveSettings = document.getElementById("saveSettings");
    
    const evalInputs = {
        bsc: document.getElementById("eval_bsc"),
        msc: document.getElementById("eval_msc"),
        research: document.getElementById("eval_research"),
        prof: document.getElementById("eval_prof"),
        english: document.getElementById("eval_english"),
        cv: document.getElementById("eval_cv")
    };
    const evalTotalScore = document.getElementById("evalTotalScore");

    Object.keys(weights).forEach(k => {
        const el = document.getElementById('w_' + k);
        if(el) el.value = weights[k];
    });

    if (settingsBtn) settingsBtn.addEventListener('click', () => settingsModal.classList.remove('hidden'));
    if (closeSettings) closeSettings.addEventListener('click', () => settingsModal.classList.add('hidden'));
    if (saveSettings) saveSettings.addEventListener('click', () => {
        let total = 0;
        Object.keys(weights).forEach(k => {
            const el = document.getElementById('w_' + k);
            if(el) {
                weights[k] = parseFloat(el.value) || 0;
                total += weights[k];
            }
        });
        document.getElementById('weightTotal').textContent = `Total: ${total}%`;
        if (total !== 100) {
            alert("Weights must total exactly 100%.");
            return;
        }
        localStorage.setItem('evalWeights', JSON.stringify(weights));
        settingsModal.classList.add('hidden');
        if (currentApplicantEmail) {
            updateScore(currentApplicantEmail);
            updateApplicantList();
        }
    });

    const resetSettingsBtn = document.getElementById("resetSettingsBtn");
    if (resetSettingsBtn) {
        resetSettingsBtn.addEventListener('click', () => {
            const defaultWeights = { bsc: 4, msc: 32, research: 16, prof: 8, english: 8, cv: 32 };
            Object.keys(defaultWeights).forEach(k => {
                const el = document.getElementById('w_' + k);
                if (el) el.value = defaultWeights[k];
            });
            document.getElementById('weightTotal').textContent = `Total: 100%`;
        });
    }

    const exportGradesBtn = document.getElementById("exportGradesBtn");
    if (exportGradesBtn) {
        exportGradesBtn.addEventListener('click', () => {
            if (Object.keys(ratings).length === 0) {
                alert("No grades to export yet.");
                return;
            }
            // Collect all emails that have either ratings or notes
            const allEmails = new Set([...Object.keys(ratings), ...Object.keys(reviewerNotes)]);
            const exportData = [];
            allEmails.forEach(email => {
                const r = ratings[email] || {};
                exportData.push({
                    "Email": email,
                    "BSc Grade": r.bsc || 0,
                    "MSc Grade": r.msc || 0,
                    "Research Exp.": r.research || 0,
                    "Prof. Exp.": r.prof || 0,
                    "English Skills": r.english || 0,
                    "CV & Cover Letter": r.cv || 0,
                    "Total Score": parseFloat(calculateScore(email).toFixed(2)),
                    "Reviewer Notes": reviewerNotes[email] || ''
                });
            });
            const ws = XLSX.utils.json_to_sheet(exportData);
            
            // Add actual Excel formulas for the Total Score (column H = index 7)
            for (let i = 0; i < exportData.length; i++) {
                const rowNum = i + 2; // 1 for header, 1 for 1-based index
                const formula = `B${rowNum}*(${weights.bsc}/100)+C${rowNum}*(${weights.msc}/100)+D${rowNum}*(${weights.research}/100)+E${rowNum}*(${weights.prof}/100)+F${rowNum}*(${weights.english}/100)+G${rowNum}*(${weights.cv}/100)`;
                const cellRef = XLSX.utils.encode_cell({c: 7, r: i + 1}); // H is column index 7
                
                // Keep the static calculated value as a fallback, but append the formula
                const val = parseFloat(calculateScore(exportData[i].Email).toFixed(2));
                ws[cellRef] = { t: 'n', v: val, f: formula };
            }

            // Set column width for Reviewer Notes (column I = index 8)
            if (!ws['!cols']) ws['!cols'] = [];
            ws['!cols'][8] = { wch: 40 };

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Grades");
            XLSX.writeFile(wb, `${originalFilename}_grades.xlsx`);
        });
    }

    // --- Save State (JSON) ---
    const saveStateBtn = document.getElementById("saveStateBtn");
    if (saveStateBtn) {
        saveStateBtn.addEventListener('click', () => {
            const state = {
                version: 1,
                weights: weights,
                ratings: ratings,
                reviewerNotes: reviewerNotes,
                markedCandidates: markedCandidates
            };
            const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${originalFilename}_review_state.json`;
            a.click();
            URL.revokeObjectURL(url);
        });
    }

    // --- Load State (JSON) ---
    const loadStateBtn = document.getElementById("loadStateBtn");
    const loadStateInput = document.getElementById("loadStateInput");
    if (loadStateBtn && loadStateInput) {
        loadStateBtn.addEventListener('click', () => loadStateInput.click());
        loadStateInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const state = JSON.parse(event.target.result);
                    
                    if (state.ratings) {
                        ratings = state.ratings;
                        localStorage.setItem('evalRatings', JSON.stringify(ratings));
                    }
                    if (state.reviewerNotes) {
                        reviewerNotes = state.reviewerNotes;
                        localStorage.setItem('reviewerNotes', JSON.stringify(reviewerNotes));
                    }
                    if (state.markedCandidates) {
                        markedCandidates = state.markedCandidates;
                        localStorage.setItem('markedCandidates', JSON.stringify(markedCandidates));
                    }
                    if (state.weights) {
                        weights = state.weights;
                        localStorage.setItem('evalWeights', JSON.stringify(weights));
                        Object.keys(weights).forEach(k => {
                            const el = document.getElementById('w_' + k);
                            if(el) el.value = weights[k];
                        });
                    }
                    
                    const count = Object.keys(state.ratings || {}).length;
                    alert(`Successfully loaded review state (${count} rated applicants).`);
                    
                    if (currentApplicantEmail) {
                        const userRatings = ratings[currentApplicantEmail] || {};
                        Object.keys(evalInputs).forEach(k => {
                            if(evalInputs[k]) {
                                evalInputs[k].value = userRatings[k] !== undefined ? userRatings[k] : '';
                            }
                        });
                        updateScore(currentApplicantEmail);
                        const notesEl = document.getElementById('reviewerNotes');
                        if (notesEl) notesEl.value = reviewerNotes[currentApplicantEmail] || '';
                    }
                    updateApplicantList();
                } catch (error) {
                    console.error("Error loading state:", error);
                    alert("Error parsing the state file. Please ensure it's a valid review state JSON.");
                }
            };
            reader.readAsText(file);
            e.target.value = '';
        });
    }

    // Helper to clamp a value to 0-10
    function clampScore(val) {
        if (isNaN(val)) return 0;
        return Math.max(0, Math.min(10, val));
    }

    Object.keys(evalInputs).forEach(k => {
        if(evalInputs[k]) {
            evalInputs[k].addEventListener('input', () => {
                if (!currentApplicantEmail) return;
                if (!ratings[currentApplicantEmail]) ratings[currentApplicantEmail] = {};
                
                let val = clampScore(parseInt(evalInputs[k].value, 10));
                
                ratings[currentApplicantEmail][k] = val;
                localStorage.setItem('evalRatings', JSON.stringify(ratings));
                updateScore(currentApplicantEmail);
                
                const activeItem = document.querySelector('.applicant-item.active .applicant-score');
                if (activeItem) {
                    activeItem.textContent = calculateScore(currentApplicantEmail).toFixed(0);
                    activeItem.classList.add('evaluated');
                }
            });

            // Clamp the displayed value when leaving the field
            evalInputs[k].addEventListener('blur', () => {
                if (evalInputs[k].value === '') return;
                let val = clampScore(parseInt(evalInputs[k].value, 10));
                evalInputs[k].value = val;
            });
        }
    });

    function calculateScore(email) {
        if (!ratings[email]) return 0;
        let score = 0;
        Object.keys(weights).forEach(k => {
            const s = clampScore(ratings[email][k] || 0);
            score += s * (weights[k] / 100);
        });
        return score;
    }

    function updateScore(email) {
        if(evalTotalScore) evalTotalScore.textContent = calculateScore(email).toFixed(0);
    }

    // --- File Upload & Parsing Logic ---
    excelUpload.addEventListener("change", (e) => {
        const file = e.target.files[0];
        if (!file) return;

        originalFilename = file.name.replace(/\.[^/.]+$/, "");
        uploadStatus.textContent = "Processing file...";

        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const data = new Uint8Array(event.target.result);
                // Parse workbook
                const workbook = XLSX.read(data, { type: "array" });
                
                // Assuming data is in the first sheet
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Convert to JSON and extract hyperlinks
                const range = XLSX.utils.decode_range(worksheet['!ref']);
                
                // Dynamically find the header row by looking for 'Email' or 'Last name'
                let headerRowIndex = range.s.r;
                for (let R = range.s.r; R <= Math.min(range.e.r, range.s.r + 10); ++R) {
                    let foundHeader = false;
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        const cell = worksheet[XLSX.utils.encode_cell({c: C, r: R})];
                        const val = cell ? (cell.w || cell.v).toString().trim() : '';
                        if (val === 'Email' || val === 'Last name') {
                            foundHeader = true;
                            break;
                        }
                    }
                    if (foundHeader) {
                        headerRowIndex = R;
                        break;
                    }
                }

                const headers = [];
                const seenHeaders = {};
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell = worksheet[XLSX.utils.encode_cell({c: C, r: headerRowIndex})];
                    let headerName = cell ? (cell.w || cell.v).toString().trim() : `Column_${C}`;
                    
                    if (seenHeaders[headerName]) {
                        seenHeaders[headerName]++;
                        headerName = headerName + "_" + seenHeaders[headerName];
                    } else {
                        seenHeaders[headerName] = 1;
                    }
                    headers.push(headerName);
                }

                const jsonData = [];
                for (let R = headerRowIndex + 1; R <= range.e.r; ++R) {
                    const rowObj = {};
                    let isEmpty = true;
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        const cellAddress = XLSX.utils.encode_cell({c: C, r: R});
                        const cell = worksheet[cellAddress];
                        if (cell) {
                            isEmpty = false;
                            let value = cell.w || cell.v || '';
                            if (cell.l && cell.l.Target) {
                                rowObj[headers[C]] = {
                                    value: value,
                                    link: cell.l.Target
                                };
                            } else {
                                rowObj[headers[C]] = value;
                            }
                        } else {
                            rowObj[headers[C]] = '';
                        }
                    }
                    if (!isEmpty) {
                        jsonData.push(rowObj);
                    }
                }

                if (jsonData.length === 0) {
                    uploadStatus.textContent = "Error: The uploaded spreadsheet appears to be empty.";
                    return;
                }

                applicantsData = jsonData;
                
                try {
                    localStorage.setItem('cachedApplicants', JSON.stringify(applicantsData));
                    localStorage.setItem('cachedFilename', originalFilename);
                } catch (e) {
                    console.warn("Could not cache file data:", e);
                }
                
                // Transition UI
                landingPage.classList.add("hidden");
                appContainer.classList.remove("hidden");
                
                // Render List (with default sort)
                updateApplicantList();

            } catch (error) {
                console.error("Error parsing Excel file:", error);
                uploadStatus.textContent = "Error parsing file. Please ensure it's a valid .xlsx file.";
            }
        };

        reader.onerror = () => {
            uploadStatus.textContent = "Error reading file.";
        };

        reader.readAsArrayBuffer(file);
    });

    // --- Search & Sort & Filter functionality ---
    const sortSelect = document.getElementById('sortSelect');
    const evalFilterSelect = document.getElementById('evalFilterSelect');
    const markedFilterSelect = document.getElementById('markedFilterSelect');
    
    function updateApplicantList() {
        const term = searchInput.value.toLowerCase();
        const sortVal = sortSelect ? sortSelect.value : 'original';
        const evalFilter = evalFilterSelect ? evalFilterSelect.value : 'all';
        const markedFilter = markedFilterSelect ? markedFilterSelect.value : 'all';
        
        let filtered = applicantsData.filter(app => {
            const getVal = (key, altKeys = []) => {
                let entry = app[key];
                for (let i = 0; i < altKeys.length && entry === undefined; i++) {
                    entry = app[altKeys[i]];
                }
                return entry && typeof entry === 'object' ? entry.value : entry;
            };
            const firstName = (getVal('General Information First name') || '').toLowerCase();
            const lastName = (getVal('Last name') || '').toLowerCase();
            const matchesSearch = firstName.includes(term) || lastName.includes(term);
            if (!matchesSearch) return false;
            
            // Evaluation filter
            const email = getVal('Email');
            const hasRatings = ratings[email] && Object.keys(ratings[email]).length > 0;
            if (evalFilter === 'evaluated' && !hasRatings) return false;
            if (evalFilter === 'notEvaluated' && hasRatings) return false;
            
            // Marked filter
            const isMarked = !!markedCandidates[email];
            if (markedFilter === 'marked' && !isMarked) return false;
            if (markedFilter === 'notMarked' && isMarked) return false;
            
            return true;
        });

        if (sortVal !== 'original') {
            filtered.sort((a, b) => {
                const getVal = (app, key) => { 
                    const e = app[key]; 
                    return e && typeof e === 'object' ? e.value : e; 
                };
                
                if (sortVal === 'nameAsc') {
                    const nameA = (getVal(a, 'Last name') + ' ' + getVal(a, 'General Information First name')).toLowerCase();
                    const nameB = (getVal(b, 'Last name') + ' ' + getVal(b, 'General Information First name')).toLowerCase();
                    return nameA.localeCompare(nameB);
                } else if (sortVal === 'scoreDesc' || sortVal === 'scoreAsc') {
                    const scoreA = calculateScore(getVal(a, 'Email'));
                    const scoreB = calculateScore(getVal(b, 'Email'));
                    return sortVal === 'scoreDesc' ? scoreB - scoreA : scoreA - scoreB;
                }
                return 0;
            });
        }
        
        renderApplicantList(filtered);
    }

    searchInput.addEventListener('input', updateApplicantList);
    if (sortSelect) sortSelect.addEventListener('change', updateApplicantList);
    if (evalFilterSelect) evalFilterSelect.addEventListener('change', updateApplicantList);
    if (markedFilterSelect) markedFilterSelect.addEventListener('change', updateApplicantList);

    // --- Render Logic ---
    function renderApplicantList(data) {
        applicantList.innerHTML = '';
        data.forEach((applicant, index) => {
            const getVal = (key) => {
                const entry = applicant[key];
                return entry && typeof entry === 'object' ? entry.value : entry;
            };
            const li = document.createElement('li');
            li.className = 'applicant-item';
            
            const firstName = getVal('General Information First name') || 'Unknown';
            const lastName = getVal('Last name') || '';
            const country = getVal('Country where you currently live') || '';
            
            const email = getVal('Email');
            const score = calculateScore(email).toFixed(0);
            const hasRatings = ratings[email] && Object.keys(ratings[email]).length > 0;
            const statusClass = hasRatings ? 'evaluated' : '';
            const isMarked = !!markedCandidates[email];
            const bookmarkIcon = isMarked ? '<span class="sidebar-bookmark" title="Marked"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="currentColor" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21l-7-5-7 5V5a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2z"></path></svg></span>' : '';
            li.innerHTML = `
                <span class="applicant-score ${statusClass}">${score}</span>
                ${bookmarkIcon}
                <h4>${firstName} ${lastName}</h4>
                <p>${country}</p>
            `;
            
            li.addEventListener('click', () => {
                // Remove active class from all
                document.querySelectorAll('.applicant-item').forEach(el => el.classList.remove('active'));
                li.classList.add('active');
                showApplicantDetails(applicant);
            });
            
            applicantList.appendChild(li);
        });
    }

    function showApplicantDetails(applicant) {
        welcomeMessage.classList.add('hidden');
        applicantDetails.classList.remove('hidden');
        
        const getVal = (key, altKeys = []) => {
            let entry = applicant[key];
            for (let i = 0; i < altKeys.length && entry === undefined; i++) {
                entry = applicant[altKeys[i]];
            }
            return entry && typeof entry === 'object' ? entry.value : entry;
        };

        const getLinkOrVal = (key, altKeys = []) => {
            let entry = applicant[key];
            for (let i = 0; i < altKeys.length && entry === undefined; i++) {
                entry = applicant[altKeys[i]];
            }
            return entry && typeof entry === 'object' ? (entry.link || entry.value) : entry;
        };

        // Helper to set text or hide parent if empty
        const setText = (id, key, altKeys = [], fallback = '-') => {
            const value = getVal(key, altKeys);
            const el = document.getElementById(id);
            if(el) el.textContent = value && value !== '/' ? value : fallback;
        };

        const setLink = (id, key) => {
            const url = getLinkOrVal(key);
            const el = document.getElementById(id);
            if(el) {
                if(url && url !== '/' && url.trim() !== '') {
                    el.href = url;
                    el.classList.remove('hidden');
                } else {
                    el.classList.add('hidden');
                }
            }
        };

        // Header
        const firstName = getVal('General Information First name') || '';
        const lastName = getVal('Last name') || '';
        document.getElementById('detailName').textContent = `${firstName} ${lastName}`;
        
        const username = getVal('Username (Choose a username)');
        const gender = getVal('Gender');
        const pref = getVal('Ranked preference for this PhD position (if you apply to other PhD positions) (1= most preferred, 3 less preferred) If you only apply to 1 PhD, please tick "Ranked 1"');
        
        const usernameEl = document.getElementById('detailUsername');
        const genderEl = document.getElementById('detailGender');
        const prefEl = document.getElementById('detailPreference');
        
        if (usernameEl) usernameEl.textContent = (username && username !== '/') ? `@${username}` : '';
        if (genderEl) genderEl.textContent = (gender && gender !== '/') ? gender : '';
        if (prefEl) prefEl.textContent = (pref && pref !== '/') ? `Preference: ${pref}` : '';
        
        currentApplicantEmail = getVal('Email');
        const userRatings = ratings[currentApplicantEmail] || {};
        Object.keys(evalInputs).forEach(k => {
            if(evalInputs[k]) {
                evalInputs[k].value = userRatings[k] !== undefined ? userRatings[k] : '';
            }
        });
        updateScore(currentApplicantEmail);
        
        // Load reviewer notes for this applicant
        const notesEl = document.getElementById('reviewerNotes');
        if (notesEl) {
            notesEl.value = reviewerNotes[currentApplicantEmail] || '';
            // Remove old listener by cloning
            const newNotesEl = notesEl.cloneNode(true);
            notesEl.parentNode.replaceChild(newNotesEl, notesEl);
            newNotesEl.addEventListener('input', () => {
                if (!currentApplicantEmail) return;
                reviewerNotes[currentApplicantEmail] = newNotesEl.value;
                localStorage.setItem('reviewerNotes', JSON.stringify(reviewerNotes));
            });
        }
        
        // Update mark toggle button
        const markToggleBtn = document.getElementById('markToggleBtn');
        if (markToggleBtn) {
            const updateMarkBtn = () => {
                const isMarked = !!markedCandidates[currentApplicantEmail];
                markToggleBtn.classList.toggle('marked', isMarked);
                markToggleBtn.title = isMarked ? 'Unmark this candidate' : 'Mark this candidate';
            };
            updateMarkBtn();
            // Remove old listener by cloning
            const newBtn = markToggleBtn.cloneNode(true);
            markToggleBtn.parentNode.replaceChild(newBtn, markToggleBtn);
            newBtn.addEventListener('click', () => {
                if (!currentApplicantEmail) return;
                if (markedCandidates[currentApplicantEmail]) {
                    delete markedCandidates[currentApplicantEmail];
                } else {
                    markedCandidates[currentApplicantEmail] = true;
                }
                localStorage.setItem('markedCandidates', JSON.stringify(markedCandidates));
                const isMarked = !!markedCandidates[currentApplicantEmail];
                newBtn.classList.toggle('marked', isMarked);
                newBtn.title = isMarked ? 'Unmark this candidate' : 'Mark this candidate';
                updateApplicantList();
            });
        }
        
        const city = getVal('City where you currently live') || '';
        const country = getVal('Country where you currently live') || '';
        document.getElementById('detailLocation').textContent = `${city}, ${country}`;

        const email = getVal('Email');
        if(email && email !== '/') {
            const emailLink = document.getElementById('linkEmail');
            emailLink.href = `mailto:${email}`;
            emailLink.classList.remove('hidden');
        } else {
            document.getElementById('linkEmail').classList.add('hidden');
        }

        setLink('linkWebsite', 'Link to your website (if available)');
        setLink('linkLinkedIn', 'Link to your LinkedIn profile (if available)');
        setLink('linkOrcid', 'ORCID (if available)');

        // Helper for University Ranking Search Link
        const setUniversityWithRankingLink = (id, key, altKeys = []) => {
            const el = document.getElementById(id);
            if (el) {
                const val = getVal(key, altKeys);
                if (val && val !== '/' && val.trim() !== '') {
                    const query = encodeURIComponent(`site:timeshighereducation.com ${val}`);
                    const searchUrl = `https://www.google.com/search?q=${query}`;
                    el.innerHTML = `${val} <a href="${searchUrl}" target="_blank" class="ranking-link" title="Search THE Ranking on Google" style="margin-left: 6px; color: var(--accent-color); font-size: 0.85em; text-decoration: none;">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align: text-bottom;"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
                    </a>`;
                } else {
                    el.textContent = 'Not specified';
                }
            }
        };

        // Education - BSc
        setText('detailBscCountry', 'Education: Bachelor Country');
        setText('detailBscYear', 'Year of graduation');
        setUniversityWithRankingLink('detailBscUni', 'University (full name)');
        setText('detailBscProgram', 'Study Program name');
        setText('detailBscGrade', 'Final Grade (in the format your_grade / maximum_grade)');

        // Education - MSc
        setText('detailMscCountry', 'Education: Master Country');
        setText('detailMscYear', 'Year');
        setUniversityWithRankingLink('detailMscUni', 'University (full name)_2', ['University (full name)2']);
        setText('detailMscProgram', 'Study Program name_2', ['Study Program name3']);
        setText('detailMscGrade', 'Final Grade (in the format your_grade / maximum_grade)_2', ['Final Grade (in the format your_grade / maximum_grade)4']);

        // Thesis
        setText('detailThesisTitle', 'Education: Master Thesis Title of your Master Thesis');
        setText('detailThesisSupervisors', 'Name of official supervisor(s) of Master Thesis (separated by comma)');
        setText('detailThesisSupervisorEmails', 'Email(s) of official supervisor(s) of Master Thesis (separated by comma)');
        
        const thesisUrl = getLinkOrVal('If the PDF of your Master Thesis is bigger than 1Mb, please provide an URL to download it');
        const thesisUrlContainer = document.getElementById('detailThesisUrlContainer');
        if (thesisUrlContainer) {
            if (thesisUrl && thesisUrl !== '/' && thesisUrl.trim() !== '') {
                document.getElementById('detailThesisUrl').href = thesisUrl;
                thesisUrlContainer.classList.remove('hidden');
            } else {
                thesisUrlContainer.classList.add('hidden');
            }
        }
        
        // References
        setText('detailReferences', 'email of reference persons');

        // Skills & Interests
        const simplifySkillLevel = (level) => {
            if (!level || level === '/') return 'Not specified';
            const str = level.toString().toLowerCase();
            if (str.includes('professional')) return 'Professional';
            if (str.includes('advanced')) return 'Advanced';
            if (str.includes('competent')) return 'Competent';
            if (str.includes('novice')) return 'Novice';
            if (str.includes('no expertise') || str.includes('none')) return 'No Expertise';
            return level;
        };

        const getSkillLevelClass = (text) => {
            switch(text) {
                case 'Professional': return 'level-professional';
                case 'Advanced': return 'level-advanced';
                case 'Competent': return 'level-competent';
                case 'Novice': return 'level-novice';
                case 'No Expertise': return 'level-none';
                default: return '';
            }
        };

        const interestsList = document.getElementById('detailInterests');
        interestsList.innerHTML = '';
        const addSkill = (nameKey, levelKey, listEl) => {
            const name = getVal(nameKey);
            const rawLevel = getVal(levelKey);
            if(name && name !== '/' && name.trim() !== '') {
                const text = simplifySkillLevel(rawLevel);
                const cssClass = getSkillLevelClass(text);
                const li = document.createElement('li');
                li.innerHTML = `<span class="skill-name">${name}</span><span class="skill-level ${cssClass}">${text}</span>`;
                listEl.appendChild(li);
            }
        };

        addSkill('Your Research Interests List your 3 primary research interests (Skill 1, Skill2 and Skill 3) Skill 1', 'Rate the level of your Skill 1 using the scale provided', interestsList);
        addSkill('Skill 2', 'Rate the level of your Skill 2 using the scale provided', interestsList);
        addSkill('Skill 3', 'Rate the level of your Skill 3 using the scale provided', interestsList);

        const techSkillsList = document.getElementById('detailTechSkills');
        techSkillsList.innerHTML = '';
        const addTechSkill = (label, levelKey) => {
            const rawLevel = getVal(levelKey);
            if(rawLevel && rawLevel !== '/' && rawLevel.trim() !== '') {
                const text = simplifySkillLevel(rawLevel);
                const cssClass = getSkillLevelClass(text);
                const li = document.createElement('li');
                li.innerHTML = `<span class="skill-name">${label}</span><span class="skill-level ${cssClass}">${text}</span>`;
                techSkillsList.appendChild(li);
            }
        };

        addTechSkill('Python', 'Main required skills for the PhD position Rate your level of expertise in Python using the scale provided');
        addTechSkill('Machine Learning (Multimodal)', 'Rate your level of expertise in Machine learning, including multimodal data analysis using the scale provided');
        addTechSkill('Explainable AI', 'Rate your level of expertise in Explainable AI using the scale provided');
        addTechSkill('Sensing & Robotics', 'Rate your level of expertise in Sensing and Robotic systems using the scale provided');
        addTechSkill('Agricultural Sustainability', 'Rate your level of expertise in Agricultural sustainability or Crop production using the scale provided');

        // Languages
        setText('detailLanguages', 'Language Languages spoken and written (separated by comma)');
        setText('detailEnglishLevel', 'English level');

        // Comments
        const comments = getVal('Optional comments (please provide additional comments regarding your expertise that fits to the PhD topic)');
        const commentsEl = document.getElementById('detailComments');
        if(comments && comments !== '/' && comments.trim() !== '') {
            commentsEl.textContent = comments;
            commentsEl.parentElement.classList.remove('hidden');
        } else {
            commentsEl.parentElement.classList.add('hidden');
        }

        // Documents
        const docsContainer = document.getElementById('detailDocuments');
        docsContainer.innerHTML = '';
        
        const docIcon = `<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>`;

        const addDoc = (label, fileKey) => {
            const entry = applicant[fileKey];
            if(!entry) return;
            const value = typeof entry === 'object' ? entry.value : entry;
            const link = typeof entry === 'object' ? entry.link : null;
            
            if(value && value !== '/' && value.trim() !== '') {
                const files = value.split(';').map(f => f.trim());
                files.forEach((file, idx) => {
                    if(!file) return;
                    const a = document.createElement('a');
                    a.className = 'doc-btn';
                    
                    if(link) {
                        a.href = 'data/' + link + '/' + file;
                    } else {
                        a.href = file;
                    }
                    a.target = '_blank';
                    
                    const displayLabel = files.length > 1 ? `${label} (${idx+1})` : label;
                    a.innerHTML = `${docIcon} <span title="${file}">${displayLabel}</span>`;
                    docsContainer.appendChild(a);
                });
            }
        };

        addDoc('Master Thesis PDF', 'Please upload the PDF of your Master Thesis');
        addDoc('Co-authored Papers', 'Upload papers you are co-author (upload in pdf format)');
        addDoc('BSc/MSc Degrees', 'a certified copy and official translation in English of bachelor and master degree');
        addDoc('Transcripts', 'a certified copy and official translation in English of the transcripts of study results for bachelor and master');
        addDoc('Grading System', 'explanation of the grading system (in English)');
        addDoc('English Proof', 'the certificate or proof of English proficiency');
        addDoc('Curriculum Vitae', 'a comprehensive curriculum vitae (in English)');
        addDoc('Cover Letter', 'a cover letter explaining motivation to join GreenFieldData (in English)');
        addDoc('Mobility Rule Proof', 'Documents that attest that you meet the eligibility criteria (mobility rule)');

        if(docsContainer.children.length === 0) {
            docsContainer.innerHTML = '<p class="text-muted">No documents uploaded.</p>';
        }

        // External Docs link
        const extLink = getLinkOrVal('If PDF of required documents are bigger than 1Mb, please provide an URL to download it');
        const extDocsContainer = document.getElementById('detailExternalDocsContainer');
        const extDocsLink = document.getElementById('detailExternalDocs');
        
        if(extLink && extLink !== '/' && extLink.trim() !== '' && extLink.startsWith('http')) {
            extDocsLink.href = extLink;
            extDocsContainer.classList.remove('hidden');
        } else {
            extDocsContainer.classList.add('hidden');
        }
    }

    // --- Close Document Logic ---
    const clearDataBtn = document.getElementById("clearDataBtn");
    if (clearDataBtn) {
        clearDataBtn.addEventListener('click', () => {
            localStorage.removeItem('cachedApplicants');
            localStorage.removeItem('cachedFilename');
            applicantsData = [];
            applicantList.innerHTML = '';
            applicantDetails.classList.add('hidden');
            welcomeMessage.classList.remove('hidden');
            
            appContainer.classList.add("hidden");
            landingPage.classList.remove("hidden");
            excelUpload.value = '';
            uploadStatus.textContent = '';
            currentApplicantEmail = null;
            
            // Reset filters
            if (evalFilterSelect) evalFilterSelect.value = 'all';
            if (markedFilterSelect) markedFilterSelect.value = 'all';
        });
    }

    // --- Auto-load cached file ---
    const cachedFile = localStorage.getItem('cachedApplicants');
    if (cachedFile) {
        try {
            applicantsData = JSON.parse(cachedFile);
            originalFilename = localStorage.getItem('cachedFilename') || 'applicants';
            
            landingPage.classList.add("hidden");
            appContainer.classList.remove("hidden");
            updateApplicantList();
        } catch (e) {
            console.error("Error loading cached file data:", e);
        }
    }
});
