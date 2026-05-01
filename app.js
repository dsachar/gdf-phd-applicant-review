// --- IndexedDB helpers for persisting the directory handle ---
const DB_NAME = 'phdReviewDB';
const DB_STORE = 'handles';
const DIR_HANDLE_KEY = 'surveyDirHandle';

function openHandleDB() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(DB_NAME, 1);
        req.onupgradeneeded = () => req.result.createObjectStore(DB_STORE);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}

async function saveDirHandle(handle) {
    const db = await openHandleDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(DB_STORE, 'readwrite');
        tx.objectStore(DB_STORE).put(handle, DIR_HANDLE_KEY);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
    });
}

async function loadDirHandle() {
    const db = await openHandleDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(DB_STORE, 'readonly');
        const req = tx.objectStore(DB_STORE).get(DIR_HANDLE_KEY);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => reject(req.error);
    });
}

async function clearDirHandle() {
    const db = await openHandleDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(DB_STORE, 'readwrite');
        tx.objectStore(DB_STORE).delete(DIR_HANDLE_KEY);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
    });
}

// Recursively iterate a directory handle and populate fileMap
async function buildFileMapFromHandle(dirHandle, prefix) {
    for await (const entry of dirHandle.values()) {
        const path = prefix ? prefix + '/' + entry.name : entry.name;
        if (entry.kind === 'file') {
            const file = await entry.getFile();
            fileMap.set(path, file);
        } else if (entry.kind === 'directory') {
            await buildFileMapFromHandle(entry, path);
        }
    }
}

let fileMap = new Map();

document.addEventListener("DOMContentLoaded", () => {
    // UI Elements
    const landingPage = document.getElementById("landingPage");
    const appContainer = document.getElementById("appContainer");
    const folderUpload = document.getElementById("folderUpload");
    const folderUploadLabel = document.getElementById("folderUploadLabel");
    const discoveryPanel = document.getElementById("discoveryPanel");
    const discoveryPlaceholder = document.getElementById("discoveryPlaceholder");
    const discoveredFolderName = document.getElementById("discoveredFolderName");
    const excelSelect = document.getElementById("excelSelect");
    const discoveryStats = document.getElementById("discoveryStats");
    const startEvaluationBtn = document.getElementById("startEvaluationBtn");
    const uploadStatus = document.getElementById("uploadStatus");
    const sidebarStats = document.getElementById("sidebarStats");

    const applicantList = document.getElementById("applicantList");
    const searchInput = document.getElementById("searchInput");
    const welcomeMessage = document.getElementById("welcomeMessage");
    const applicantDetails = document.getElementById("applicantDetails");
    
    let applicantsData = [];
    let originalFilename = 'applicants';
    let excelFiles = [];
    let storedDirHandle = null; // Persisted directory handle (File System Access API)

    function updateTopBarTitle(name) {
        const titleEl = document.getElementById('topBarTitle');
        if (titleEl) titleEl.textContent = name;
    }

    // Rating state
    let currentApplicantEmail = null;
    let weights = JSON.parse(localStorage.getItem('evalWeights')) || { bsc: 4, msc: 32, research: 16, prof: 8, english: 8, cv: 32 };
    let ratings = JSON.parse(localStorage.getItem('evalRatings')) || {};
    let markedCandidates = JSON.parse(localStorage.getItem('markedCandidates')) || {};
    let reviewerNotes = JSON.parse(localStorage.getItem('reviewerNotes')) || {};
    let docBasePath = localStorage.getItem('docBasePath') || 'data/';
    let secondaryReviewers = JSON.parse(localStorage.getItem('secondaryReviewers')) || {};
    
    let primaryReviewerName = localStorage.getItem('primaryReviewerName');
    let consensusRatings = JSON.parse(localStorage.getItem('consensusRatings')) || {};
    let consensusNotes = JSON.parse(localStorage.getItem('consensusNotes')) || {};

    if (!primaryReviewerName) {
        primaryReviewerName = prompt("Welcome! Please enter your name to identify your reviews:", "Reviewer") || "Reviewer";
        localStorage.setItem('primaryReviewerName', primaryReviewerName);
    }

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

    if (settingsBtn) settingsBtn.addEventListener('click', () => {
        const pathInput = document.getElementById('doc_base_path');
        if (pathInput) pathInput.value = docBasePath;
        settingsModal.classList.remove('hidden');
    });
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
        const pathInput = document.getElementById('doc_base_path');
        if (pathInput) docBasePath = pathInput.value || 'data/';
        localStorage.setItem('docBasePath', docBasePath);

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

    function doExportGrades(useConsensus) {
        const targetRatings = useConsensus ? consensusRatings : ratings;
        const targetNotes = useConsensus ? consensusNotes : reviewerNotes;

        if (Object.keys(targetRatings).length === 0 && Object.keys(targetNotes).length === 0) {
            alert(`No ${useConsensus ? 'consensus ' : ''}grades to export yet.`);
            return;
        }
        // Collect all emails that have either ratings or notes
        const allEmails = new Set([...Object.keys(targetRatings), ...Object.keys(targetNotes)]);
        const exportData = [];
        allEmails.forEach(email => {
            const r = targetRatings[email] || {};
            // Look up applicant record to get name fields
            const applicant = applicantsData.find(a => {
                const e = a['Email'];
                return (e && typeof e === 'object' ? e.value : e) === email;
            });
            const getAppVal = (key) => {
                if (!applicant) return '';
                const entry = applicant[key];
                return entry && typeof entry === 'object' ? entry.value : (entry || '');
            };
            
            const score = calculateSpecificScore(email, useConsensus ? 'consensus' : 'primary');
            
            exportData.push({
                "Email": email,
                "First Name": getAppVal('General Information First name'),
                "Last Name": getAppVal('Last name'),
                "BSc Grade": r.bsc || 0,
                "MSc Grade": r.msc || 0,
                "Research Exp.": r.research || 0,
                "Prof. Exp.": r.prof || 0,
                "English Skills": r.english || 0,
                "CV & Cover Letter": r.cv || 0,
                "Total Score": parseFloat(score.toFixed(2)),
                "Reviewer Notes": targetNotes[email] || ''
            });
        });
        const ws = XLSX.utils.json_to_sheet(exportData);
        
        // Add actual Excel formulas for the Total Score (column J = index 9)
        for (let i = 0; i < exportData.length; i++) {
            const rowNum = i + 2; // 1 for header, 1 for 1-based index
            const formula = `D${rowNum}*(${weights.bsc}/100)+E${rowNum}*(${weights.msc}/100)+F${rowNum}*(${weights.research}/100)+G${rowNum}*(${weights.prof}/100)+H${rowNum}*(${weights.english}/100)+I${rowNum}*(${weights.cv}/100)`;
            const cellRef = XLSX.utils.encode_cell({c: 9, r: i + 1}); // J is column index 9
            
            // Keep the static calculated value as a fallback, but append the formula
            const score = calculateSpecificScore(exportData[i].Email, useConsensus ? 'consensus' : 'primary');
            const val = parseFloat(score.toFixed(2));
            ws[cellRef] = { t: 'n', v: val, f: formula };
        }

        // Set column width for Reviewer Notes (column K = index 10)
        if (!ws['!cols']) ws['!cols'] = [];
        ws['!cols'][10] = { wch: 40 };

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, useConsensus ? "Consensus Grades" : "Grades");
        const filenameSuffix = useConsensus ? "_consensus_grades.xlsx" : "_grades.xlsx";
        XLSX.writeFile(wb, `${originalFilename}${filenameSuffix}`);
    }

    const exportGradesBtn = document.getElementById("exportGradesBtn");
    const exportModal = document.getElementById("exportModal");
    const closeExportModal = document.getElementById("closeExportModal");
    const exportPersonalBtn = document.getElementById("exportPersonalBtn");
    const exportConsensusBtn = document.getElementById("exportConsensusBtn");

    if (exportGradesBtn) {
        exportGradesBtn.addEventListener('click', () => {
            if (Object.keys(secondaryReviewers).length > 0) {
                // Show modal with two options
                exportModal.classList.remove('hidden');
            } else {
                // No secondary reviewers, export personal directly
                doExportGrades(false);
            }
        });
    }
    if (closeExportModal) {
        closeExportModal.addEventListener('click', () => exportModal.classList.add('hidden'));
    }
    if (exportPersonalBtn) {
        exportPersonalBtn.addEventListener('click', () => {
            exportModal.classList.add('hidden');
            doExportGrades(false);
        });
    }
    if (exportConsensusBtn) {
        exportConsensusBtn.addEventListener('click', () => {
            exportModal.classList.add('hidden');
            doExportGrades(true);
        });
    }

    // --- Save State (JSON) ---
    const saveStateBtn = document.getElementById("saveStateBtn");
    const exportStateModal = document.getElementById("exportStateModal");
    const closeExportStateModal = document.getElementById("closeExportStateModal");
    const exportStateButtons = document.getElementById("exportStateButtons");

    function doExportState(exportRatings, exportNotes, exportMarked, label) {
        const state = {
            version: 1,
            weights: weights,
            ratings: exportRatings,
            reviewerNotes: exportNotes,
            markedCandidates: exportMarked || {},
            docBasePath: docBasePath,
            primaryReviewerName: label
        };
        const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const safeName = label.replace(/[^a-zA-Z0-9]/g, '_');
        a.download = `${originalFilename}_${safeName}_review_state.json`;
        a.click();
        URL.revokeObjectURL(url);
    }

    if (saveStateBtn) {
        saveStateBtn.addEventListener('click', () => {
            if (Object.keys(secondaryReviewers).length > 0) {
                // Build dynamic buttons
                exportStateButtons.innerHTML = '';
                
                // Personal state button
                const personalBtn = document.createElement('button');
                personalBtn.className = 'btn-primary';
                personalBtn.style.cssText = 'width: 100%; padding: 14px; font-size: 0.95rem; display: flex; align-items: center; justify-content: center; gap: 10px;';
                personalBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg><span>${primaryReviewerName || 'Personal'} (Your State)</span>`;
                personalBtn.addEventListener('click', () => {
                    exportStateModal.classList.add('hidden');
                    doExportState(ratings, reviewerNotes, markedCandidates, primaryReviewerName || 'Personal');
                });
                exportStateButtons.appendChild(personalBtn);
                
                // Secondary reviewer buttons
                for (const [name, data] of Object.entries(secondaryReviewers)) {
                    const secBtn = document.createElement('button');
                    secBtn.className = 'btn-primary';
                    secBtn.style.cssText = 'width: 100%; padding: 14px; font-size: 0.95rem; display: flex; align-items: center; justify-content: center; gap: 10px; background: #6b7280;';
                    secBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg><span>${name}</span>`;
                    secBtn.addEventListener('click', () => {
                        exportStateModal.classList.add('hidden');
                        doExportState(data.ratings || {}, data.notes || {}, {}, name);
                    });
                    exportStateButtons.appendChild(secBtn);
                }
                
                // Consensus button
                const conBtn = document.createElement('button');
                conBtn.className = 'btn-primary';
                conBtn.style.cssText = 'width: 100%; padding: 14px; font-size: 0.95rem; display: flex; align-items: center; justify-content: center; gap: 10px; background: #2563eb;';
                conBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg><span>Consensus</span>`;
                conBtn.addEventListener('click', () => {
                    exportStateModal.classList.add('hidden');
                    doExportState(consensusRatings, consensusNotes, markedCandidates, 'Consensus');
                });
                exportStateButtons.appendChild(conBtn);
                
                exportStateModal.classList.remove('hidden');
            } else {
                // No secondary reviewers, export directly
                doExportState(ratings, reviewerNotes, markedCandidates, primaryReviewerName || 'Personal');
            }
        });
    }
    if (closeExportStateModal) {
        closeExportStateModal.addEventListener('click', () => exportStateModal.classList.add('hidden'));
    }

    // --- Load State (JSON) with Multiple Reviewers Support ---
    const loadStateBtn = document.getElementById("loadStateBtn");
    const loadStateInput = document.getElementById("loadStateInput");
    
    // Modal elements
    const importModal = document.getElementById("importModal");
    const closeImportModal = document.getElementById("closeImportModal");
    const confirmImportBtn = document.getElementById("confirmImportBtn");
    const importReviewerName = document.getElementById("importReviewerName");
    const importModeRadios = document.getElementsByName("importMode");
    
    let pendingImportState = null;

    if (importModeRadios) {
        importModeRadios.forEach(radio => {
            radio.addEventListener('change', (e) => {
                if (importReviewerName) {
                    importReviewerName.disabled = e.target.value !== 'add';
                    if (e.target.value === 'add') importReviewerName.focus();
                }
            });
        });
    }

    const closeImport = () => {
        if (importModal) importModal.classList.add('hidden');
        pendingImportState = null;
        if (loadStateInput) loadStateInput.value = '';
        if (importReviewerName) importReviewerName.value = '';
    };

    if (closeImportModal) closeImportModal.addEventListener('click', closeImport);

    // Simple hash function for deduplication
    function hashRatings(r) {
        return JSON.stringify(r || {});
    }

    if (loadStateBtn && loadStateInput) {
        loadStateBtn.addEventListener('click', () => loadStateInput.click());
        loadStateInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const state = JSON.parse(event.target.result);
                    
                    // Check for duplicates
                    const importedHash = hashRatings(state.ratings);
                    const currentHash = hashRatings(ratings);
                    
                    if (importedHash === currentHash) {
                        alert("This file appears to be identical to your current state. Import cancelled.");
                        loadStateInput.value = '';
                        return;
                    }
                    
                    let duplicateReviewer = null;
                    for (const [name, rev] of Object.entries(secondaryReviewers)) {
                        if (hashRatings(rev.ratings) === importedHash) {
                            duplicateReviewer = name;
                            break;
                        }
                    }
                    
                    if (duplicateReviewer) {
                        alert(`This file appears to be identical to the imported state for "${duplicateReviewer}". Import cancelled.`);
                        loadStateInput.value = '';
                        return;
                    }

                    // If we reach here, it's a new state. Show the modal.
                    pendingImportState = state;
                    if (importModal) importModal.classList.remove('hidden');

                } catch (error) {
                    console.error("Error parsing JSON:", error);
                    alert("Invalid review state file.");
                    loadStateInput.value = '';
                }
            };
            reader.readAsText(file);
        });
    }

    if (confirmImportBtn) {
        confirmImportBtn.addEventListener('click', () => {
            if (!pendingImportState) return;
            
            const mode = Array.from(importModeRadios).find(r => r.checked)?.value;
            const state = pendingImportState;

            if (mode === 'override') {
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
                if (state.docBasePath) {
                    docBasePath = state.docBasePath;
                    localStorage.setItem('docBasePath', docBasePath);
                }
                if (state.primaryReviewerName) {
                    primaryReviewerName = state.primaryReviewerName;
                    localStorage.setItem('primaryReviewerName', primaryReviewerName);
                }
                if (state.consensusRatings) {
                    consensusRatings = state.consensusRatings;
                    localStorage.setItem('consensusRatings', JSON.stringify(consensusRatings));
                }
                if (state.consensusNotes) {
                    consensusNotes = state.consensusNotes;
                    localStorage.setItem('consensusNotes', JSON.stringify(consensusNotes));
                }
                
                const count = Object.keys(state.ratings || {}).length;
                alert(`Successfully overridden review state (${count} rated applicants).`);
                
            } else if (mode === 'add') {
                let name = importReviewerName.value.trim();
                if (!name) {
                    // Generate a default name
                    const count = Object.keys(secondaryReviewers).length + 2;
                    name = `Reviewer ${count}`;
                }
                
                if (secondaryReviewers[name]) {
                    alert(`A reviewer named "${name}" already exists. Please choose a different name.`);
                    return; // keep modal open
                }

                secondaryReviewers[name] = {
                    ratings: state.ratings || {},
                    notes: state.reviewerNotes || {}
                };
                localStorage.setItem('secondaryReviewers', JSON.stringify(secondaryReviewers));
                if (typeof updateSortOptions === 'function') updateSortOptions();
                
                alert(`Successfully added "${name}" as a secondary reviewer.`);
            }

            closeImport();

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
            recalculateConsensus();
            updateApplicantList();
        });
    }

    function recalculateConsensus() {
        if (Object.keys(secondaryReviewers).length === 0) return;
        
        const allEmails = new Set([...Object.keys(ratings)]);
        for (const reviewerData of Object.values(secondaryReviewers)) {
            if (reviewerData.ratings) {
                Object.keys(reviewerData.ratings).forEach(email => allEmails.add(email));
            }
        }
        
        allEmails.forEach(email => {
            let counts = { bsc: 0, msc: 0, research: 0, prof: 0, english: 0, cv: 0 };
            let sums = { bsc: 0, msc: 0, research: 0, prof: 0, english: 0, cv: 0 };
            let combinedNotes = [];
            
            // Add primary reviewer
            if (ratings[email]) {
                Object.keys(sums).forEach(k => {
                    if (ratings[email][k] !== undefined) {
                        sums[k] += ratings[email][k];
                        counts[k]++;
                    }
                });
            }
            if (reviewerNotes[email] && reviewerNotes[email].trim()) {
                combinedNotes.push(`${primaryReviewerName}:\n${reviewerNotes[email].trim()}`);
            }
            
            // Add secondary reviewers
            for (const [reviewerName, reviewerData] of Object.entries(secondaryReviewers)) {
                if (reviewerData.ratings && reviewerData.ratings[email]) {
                    Object.keys(sums).forEach(k => {
                        if (reviewerData.ratings[email][k] !== undefined) {
                            sums[k] += reviewerData.ratings[email][k];
                            counts[k]++;
                        }
                    });
                }
                if (reviewerData.notes && reviewerData.notes[email] && reviewerData.notes[email].trim()) {
                    combinedNotes.push(`${reviewerName}:\n${reviewerData.notes[email].trim()}`);
                }
            }
            
            // Calculate average
            if (!consensusRatings[email]) consensusRatings[email] = {};
            Object.keys(sums).forEach(k => {
                if (counts[k] > 0) {
                    consensusRatings[email][k] = parseFloat((sums[k] / counts[k]).toFixed(1));
                }
            });
            
            if (!consensusNotes[email]) {
                consensusNotes[email] = combinedNotes.join('\n\n');
            }
        });
        
        localStorage.setItem('consensusRatings', JSON.stringify(consensusRatings));
        localStorage.setItem('consensusNotes', JSON.stringify(consensusNotes));
        
        // Re-render current applicant to show consensus panel
        if (currentApplicantEmail) {
            const applicant = applicantsData.find(a => {
                const e = a['Email'];
                return (e && typeof e === 'object' ? e.value : e) === currentApplicantEmail;
            });
            if (applicant) showApplicantDetails(applicant);
        }
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
                    activeItem.textContent = calculateScore(currentApplicantEmail).toFixed(1);
                    activeItem.classList.add('evaluated');
                }
                updateApplicantList();
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

    function calculateSpecificScore(email, type) {
        let revRatings = {};
        if (type === 'consensus') {
            revRatings = consensusRatings[email] || {};
        } else if (type.startsWith('sec_')) {
            const name = type.substring(4);
            if (secondaryReviewers[name] && secondaryReviewers[name].ratings) {
                revRatings = secondaryReviewers[name].ratings[email] || {};
            }
        } else {
            revRatings = ratings[email] || {};
        }
        
        let score = 0;
        Object.keys(weights).forEach(k => {
            const s = clampScore(revRatings[k] || 0);
            score += s * (weights[k] / 100);
        });
        return score;
    }

    function updateScore(email) {
        if(evalTotalScore) evalTotalScore.textContent = calculateScore(email).toFixed(1);
    }

    // --- Folder Upload & Discovery Logic ---

    // Common function to process discovered files and show the discovery panel
    function showDiscovery(rootName) {
        discoveredFolderName.textContent = rootName;

        // Separate excel files from the rest by inspecting fileMap
        excelFiles = [];
        const docEntries = new Map();
        for (const [relPath, file] of fileMap) {
            const lower = relPath.toLowerCase();
            if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
                excelFiles.push(file);
            } else {
                docEntries.set(relPath, file);
            }
        }
        // Keep only non-excel entries in fileMap
        fileMap = docEntries;

        if (excelFiles.length === 0) {
            uploadStatus.textContent = "Error: No Excel files (.xlsx or .xls) found in this folder.";
            discoveryPanel.classList.add('hidden');
            return;
        }

        // Populate Spreadsheet Selection
        excelSelect.innerHTML = '';
        excelFiles.forEach((file, idx) => {
            const option = document.createElement('option');
            option.value = idx;
            option.textContent = file.name;
            excelSelect.appendChild(option);
        });

        const excelSelectionArea = document.getElementById('excelSelectionArea');
        if (excelFiles.length === 1) {
            excelSelectionArea.classList.add('hidden');
        } else {
            excelSelectionArea.classList.remove('hidden');
        }

        discoveryStats.textContent = `Found ${excelFiles.length} spreadsheet(s) and ${fileMap.size} document(s).`;
        discoveryPanel.classList.remove('hidden');
        if (discoveryPlaceholder) discoveryPlaceholder.classList.add('hidden');
        if (folderUploadLabel) folderUploadLabel.textContent = "Select Different Survey Folder";
        uploadStatus.textContent = "";
    }

    // Try the modern File System Access API for folder selection
    async function selectFolderViaFSAPI() {
        try {
            const dirHandle = await window.showDirectoryPicker({ mode: 'read' });
            storedDirHandle = dirHandle;
            fileMap.clear();
            excelFiles = [];
            await buildFileMapFromHandle(dirHandle, '');
            await saveDirHandle(dirHandle);
            showDiscovery(dirHandle.name);
        } catch (err) {
            if (err.name !== 'AbortError') {
                console.error('showDirectoryPicker failed:', err);
            }
        }
    }

    // Wire up the folder selection button
    if (window.showDirectoryPicker) {
        // Modern API available – hide the native input and use the label as a button
        folderUpload.style.display = 'none';
        folderUploadLabel.removeAttribute('for');
        folderUploadLabel.addEventListener('click', (e) => {
            e.preventDefault();
            selectFolderViaFSAPI();
        });
    }

    // Fallback: traditional <input webkitdirectory> (also fires on non-FSAPI browsers)
    folderUpload.addEventListener("change", (e) => {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;

        fileMap.clear();
        excelFiles = [];
        storedDirHandle = null;
        clearDirHandle(); // clear any stored handle since we're using input fallback
        
        // Find the root folder name from the first file's webkitRelativePath
        const firstPath = files[0].webkitRelativePath;
        const rootFolder = firstPath.split('/')[0];

        files.forEach(file => {
            const relPath = file.webkitRelativePath.replace(rootFolder + '/', '');
            fileMap.set(relPath, file);
        });

        showDiscovery(rootFolder);
    });

    startEvaluationBtn.addEventListener('click', () => {
        const selectedIdx = excelSelect.value;
        const file = excelFiles[selectedIdx];
        if (!file) return;

        originalFilename = file.name.replace(/\.[^/.]+$/, "");
        updateTopBarTitle(originalFilename);
        uploadStatus.textContent = "Processing data...";

        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: "array" });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const range = XLSX.utils.decode_range(worksheet['!ref']);
                
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
                                rowObj[headers[C]] = { value: value, link: cell.l.Target };
                            } else {
                                rowObj[headers[C]] = value;
                            }
                        } else {
                            rowObj[headers[C]] = '';
                        }
                    }
                    if (!isEmpty) jsonData.push(rowObj);
                }

                if (jsonData.length === 0) {
                    uploadStatus.textContent = "Error: The selected spreadsheet appears to be empty.";
                    return;
                }

                applicantsData = jsonData;
                
                try {
                    localStorage.setItem('cachedApplicants', JSON.stringify(applicantsData));
                    localStorage.setItem('cachedFilename', originalFilename);
                } catch (e) {}
                
                landingPage.classList.add("hidden");
                appContainer.classList.remove("hidden");
                updateApplicantList();

            } catch (error) {
                console.error("Error parsing Excel file:", error);
                uploadStatus.textContent = "Error parsing file. Please ensure it's a valid .xlsx file.";
            }
        };
        reader.readAsArrayBuffer(file);
    });

    // --- Search & Sort & Filter functionality ---
    const sortSelect = document.getElementById('sortSelect');
    const evalFilterSelect = document.getElementById('evalFilterSelect');
    const markedFilterSelect = document.getElementById('markedFilterSelect');
    
    function updateSortOptions() {
        if (!sortSelect) return;
        const currentVal = sortSelect.value;
        
        sortSelect.innerHTML = `
            <option value="original">Sort by: Original Order</option>
            <option value="nameAsc">Sort by: Name (Last, First)</option>
        `;
        
        const hasSecondary = Object.keys(secondaryReviewers).length > 0;
        
        if (hasSecondary) {
            sortSelect.innerHTML += `
                <option value="scoreDesc_primary">Rating: ${primaryReviewerName} (Highest)</option>
            `;
            for (const reviewerName of Object.keys(secondaryReviewers)) {
                sortSelect.innerHTML += `
                    <option value="scoreDesc_sec_${reviewerName}">Rating: ${reviewerName} (Highest)</option>
                `;
            }
            sortSelect.innerHTML += `
                <option value="scoreDesc_consensus">Rating: Consensus (Highest)</option>
            `;
        } else {
            sortSelect.innerHTML += `
                <option value="scoreDesc_primary">Sort by: Rating (Highest first)</option>
            `;
        }
        
        // Restore previous value if it still exists
        const optionExists = Array.from(sortSelect.options).some(opt => opt.value === currentVal);
        if (optionExists) {
            sortSelect.value = currentVal;
        } else {
            sortSelect.value = 'original';
        }
    }
    updateSortOptions();

    
    function updateApplicantList() {
        const term = searchInput.value.toLowerCase();
        const sortVal = sortSelect ? sortSelect.value : 'original';
        const evalFilter = evalFilterSelect ? evalFilterSelect.value : 'all';
        const markedFilter = markedFilterSelect ? markedFilterSelect.value : 'all';
        
        const getVal = (app, key, altKeys = []) => {
            let entry = app[key];
            for (let i = 0; i < altKeys.length && entry === undefined; i++) {
                entry = app[altKeys[i]];
            }
            return entry && typeof entry === 'object' ? entry.value : entry;
        };

        // 1. We no longer filter by search term, but we will highlight it in the UI
        const searchFiltered = applicantsData;

        // 2. Calculate counts for Evaluation Dropdown (depends on Search + Marked Filter)
        const forEvalDropdown = searchFiltered.filter(app => {
            const email = getVal(app, 'Email');
            const isMarked = !!markedCandidates[email];
            if (markedFilter === 'marked' && !isMarked) return false;
            if (markedFilter === 'notMarked' && isMarked) return false;
            return true;
        });
        const evalCounts = { all: forEvalDropdown.length, evaluated: 0, notEvaluated: 0 };
        forEvalDropdown.forEach(app => {
            const email = getVal(app, 'Email');
            const hasRatings = ratings[email] && Object.keys(ratings[email]).length > 0;
            if (hasRatings) evalCounts.evaluated++;
            else evalCounts.notEvaluated++;
        });

        // 3. Calculate counts for Marked Dropdown (depends on Search + Eval Filter)
        const forMarkedDropdown = searchFiltered.filter(app => {
            const email = getVal(app, 'Email');
            const hasRatings = ratings[email] && Object.keys(ratings[email]).length > 0;
            if (evalFilter === 'evaluated' && !hasRatings) return false;
            if (evalFilter === 'notEvaluated' && hasRatings) return false;
            return true;
        });
        const markedCounts = { all: forMarkedDropdown.length, marked: 0, notMarked: 0 };
        forMarkedDropdown.forEach(app => {
            const email = getVal(app, 'Email');
            if (markedCandidates[email]) markedCounts.marked++;
            else markedCounts.notMarked++;
        });

        // 4. Final filtered list (Search + Eval + Marked)
        let filtered = searchFiltered.filter(app => {
            const email = getVal(app, 'Email');
            if (currentApplicantEmail && email === currentApplicantEmail) return true;
            const hasRatings = ratings[email] && Object.keys(ratings[email]).length > 0;
            const isMarked = !!markedCandidates[email];
            
            if (evalFilter === 'evaluated' && !hasRatings) return false;
            if (evalFilter === 'notEvaluated' && hasRatings) return false;
            if (markedFilter === 'marked' && !isMarked) return false;
            if (markedFilter === 'notMarked' && isMarked) return false;
            
            return true;
        });

        // Update Dropdown Labels
        if (evalFilterSelect) {
            evalFilterSelect.options[0].textContent = `Evaluation: Show All (${evalCounts.all})`;
            evalFilterSelect.options[1].textContent = `Only Evaluated (${evalCounts.evaluated})`;
            evalFilterSelect.options[2].textContent = `Only Not Evaluated (${evalCounts.notEvaluated})`;
        }
        if (markedFilterSelect) {
            markedFilterSelect.options[0].textContent = `Marked: Show All (${markedCounts.all})`;
            markedFilterSelect.options[1].textContent = `Only Marked (${markedCounts.marked})`;
            markedFilterSelect.options[2].textContent = `Only Not Marked (${markedCounts.notMarked})`;
        }

        // Update Stats
        if (sidebarStats) {
            sidebarStats.textContent = `Showing ${filtered.length} of ${applicantsData.length} applicants`;
        }

        if (sortVal !== 'original') {
            filtered.sort((a, b) => {
                if (sortVal === 'nameAsc') {
                    const nameA = (getVal(a, 'Last name') + ' ' + getVal(a, 'General Information First name')).toLowerCase();
                    const nameB = (getVal(b, 'Last name') + ' ' + getVal(b, 'General Information First name')).toLowerCase();
                    return nameA.localeCompare(nameB);
                } else if (sortVal.startsWith('scoreDesc') || sortVal.startsWith('scoreAsc')) {
                    const isDesc = sortVal.startsWith('scoreDesc');
                    const parts = sortVal.split('_');
                    let type = parts[1] || 'primary';
                    if (type === 'sec' && parts[2]) type = 'sec_' + parts[2];
                    
                    const scoreA = calculateSpecificScore(getVal(a, 'Email'), type);
                    const scoreB = calculateSpecificScore(getVal(b, 'Email'), type);
                    return isDesc ? scoreB - scoreA : scoreA - scoreB;
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
        const term = searchInput.value.toLowerCase();
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
            
            let displayFirstName = firstName;
            let displayLastName = lastName;
            
            if (term) {
                const highlight = (str) => {
                    if (!str) return str;
                    const lowerStr = str.toLowerCase();
                    const idx = lowerStr.indexOf(term);
                    if (idx >= 0) {
                        return str.substring(0, idx) + `<mark style="background-color: #fde047; padding: 0 2px; border-radius: 2px;">${str.substring(idx, idx + term.length)}</mark>` + str.substring(idx + term.length);
                    }
                    return str;
                };
                displayFirstName = highlight(firstName);
                displayLastName = highlight(lastName);
            }
            
            const email = getVal('Email');
            
            // Build scores HTML
            let scoresHtml = '';
            const hasRatings = ratings[email] && Object.keys(ratings[email]).length > 0;
            const primaryScore = calculateSpecificScore(email, 'primary').toFixed(1);
            
            if (Object.keys(secondaryReviewers).length > 0) {
                // Multiple reviewers: show all
                if (hasRatings) {
                    scoresHtml += `<span class="applicant-score primary-score evaluated" title="Your Score">${primaryScore}</span>`;
                }
                
                for (const reviewerName of Object.keys(secondaryReviewers)) {
                    const revHasRatings = secondaryReviewers[reviewerName].ratings && secondaryReviewers[reviewerName].ratings[email] && Object.keys(secondaryReviewers[reviewerName].ratings[email]).length > 0;
                    if (revHasRatings) {
                        const revScore = calculateSpecificScore(email, `sec_${reviewerName}`).toFixed(1);
                        scoresHtml += `<span class="applicant-score secondary-score evaluated" title="Score by ${reviewerName}">${revScore}</span>`;
                    }
                }
                
                if (scoresHtml === '') {
                    // No one has evaluated yet
                    scoresHtml = `<span class="applicant-score">0.0</span>`;
                }
            } else {
                // Single reviewer: standard display
                const statusClass = hasRatings ? 'evaluated' : '';
                scoresHtml = `<span class="applicant-score ${statusClass}">${primaryScore}</span>`;
            }

            const isMarked = !!markedCandidates[email];
            const bookmarkIcon = isMarked ? '<span class="sidebar-bookmark" title="Marked"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="currentColor" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21l-7-5-7 5V5a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2z"></path></svg></span>' : '';
            li.innerHTML = `
                <div class="applicant-scores-container">${scoresHtml}</div>
                ${bookmarkIcon}
                <h4>${displayFirstName} ${displayLastName}</h4>
                <p>${country}</p>
            `;
            
            if (email === currentApplicantEmail) {
                li.classList.add('active');
            }
            
            li.addEventListener('click', () => {
                // Remove active class from all
                document.querySelectorAll('.applicant-item').forEach(el => el.classList.remove('active'));
                li.classList.add('active');
                showApplicantDetails(applicant);
            });
            
            applicantList.appendChild(li);
        });
        
        if (term) {
            const firstMatch = applicantList.querySelector('mark');
            if (firstMatch) {
                firstMatch.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }
        }
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

        // Normalize a URL: ensure it has a protocol so it doesn't resolve to localhost.
        // Also converts bare ORCID identifiers to full ORCID URLs.
        const normalizeUrl = (raw) => {
            if (!raw) return raw;
            let url = raw.trim();
            // Bare ORCID identifier (e.g. "0000-0002-1234-5678")
            if (/^\d{4}-\d{4}-\d{4}-\d{3}[\dX]$/i.test(url)) {
                return `https://orcid.org/${url}`;
            }
            // Already has a protocol
            if (/^https?:\/\//i.test(url) || /^mailto:/i.test(url)) {
                return url;
            }
            // Starts with www or a known domain — prepend https://
            if (/^(www\.|linkedin\.com|orcid\.org)/i.test(url)) {
                return `https://${url}`;
            }
            // Looks like a domain path (contains a dot before any slash) — prepend https://
            if (/^[^\/]+\.[a-z]{2,}/i.test(url)) {
                return `https://${url}`;
            }
            // Fallback: return as-is (could be an unusual value)
            return url;
        };

        const setLink = (id, key) => {
            const raw = getLinkOrVal(key);
            const el = document.getElementById(id);
            if(el) {
                if(raw && raw !== '/' && raw.trim() !== '') {
                    el.href = normalizeUrl(raw);
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
        
        // Render Secondary Reviewer Panels
        const evalContainer = document.getElementById('evaluationsContainer');
        
        // Remove any existing consensus panel that might be outside evalContainer
        const existingConPanel = document.querySelector('.consensus-panel');
        if (existingConPanel) {
            existingConPanel.remove();
        }
        
        if (evalContainer) {
            // Keep the first child (primary evaluation) and remove the rest
            while (evalContainer.children.length > 1) {
                evalContainer.removeChild(evalContainer.lastChild);
            }
            
            // Loop through secondary reviewers
            for (const [reviewerName, reviewerData] of Object.entries(secondaryReviewers)) {
                const revRatings = reviewerData.ratings[currentApplicantEmail] || {};
                const revNotes = reviewerData.notes[currentApplicantEmail] || '';
                
                // Calculate their score
                let revScore = 0;
                let totalWeight = 0;
                Object.keys(weights).forEach(k => {
                    totalWeight += weights[k];
                    revScore += (revRatings[k] || 0) * weights[k];
                });
                revScore = totalWeight > 0 ? (revScore / totalWeight) : 0;
                
                // Create panel clone
                const panelHtml = `
                    <section class="card evaluation-panel secondary-reviewer-panel" style="flex: 1; min-width: 300px; margin-bottom: 0;">
                        <div class="eval-header">
                            <div style="display: flex; align-items: center; justify-content: space-between; width: 100%;">
                                <h2>${reviewerName}</h2>
                                <button class="btn-secondary remove-reviewer-btn" data-name="${reviewerName}" style="padding: 2px 6px; font-size: 0.75rem; border: 1px solid var(--border-color); border-radius: 4px; background: transparent; cursor: pointer; color: var(--text-muted);" title="Remove Reviewer">&times;</button>
                            </div>
                            <div class="total-score">Score: <strong>${revScore.toFixed(2)}</strong> / 10</div>
                        </div>
                        <div class="eval-grid">
                            <div class="eval-input"><label>BSc Grade</label><input type="text" value="${revRatings.bsc !== undefined ? revRatings.bsc : '-'}" disabled></div>
                            <div class="eval-input"><label>MSc Grade</label><input type="text" value="${revRatings.msc !== undefined ? revRatings.msc : '-'}" disabled></div>
                            <div class="eval-input"><label>Research Exp.</label><input type="text" value="${revRatings.research !== undefined ? revRatings.research : '-'}" disabled></div>
                            <div class="eval-input"><label>Prof. Exp.</label><input type="text" value="${revRatings.prof !== undefined ? revRatings.prof : '-'}" disabled></div>
                            <div class="eval-input"><label>English Skills</label><input type="text" value="${revRatings.english !== undefined ? revRatings.english : '-'}" disabled></div>
                            <div class="eval-input"><label>CV & Cover Letter</label><input type="text" value="${revRatings.cv !== undefined ? revRatings.cv : '-'}" disabled></div>
                        </div>
                        <div class="reviewer-notes-container mt-4">
                            <label class="reviewer-notes-label">Notes</label>
                            <textarea class="reviewer-notes" readonly>${revNotes}</textarea>
                        </div>
                    </section>
                `;
                
                const template = document.createElement('template');
                template.innerHTML = panelHtml.trim();
                const panelEl = template.content.firstChild;
                
                // Add event listener to remove button
                const removeBtn = panelEl.querySelector('.remove-reviewer-btn');
                if (removeBtn) {
                    removeBtn.addEventListener('click', (e) => {
                        e.stopPropagation();
                        if (confirm(`Are you sure you want to remove reviewer "${reviewerName}"?`)) {
                            delete secondaryReviewers[reviewerName];
                            localStorage.setItem('secondaryReviewers', JSON.stringify(secondaryReviewers));
                            if (typeof updateSortOptions === 'function') updateSortOptions();
                            showApplicantDetails(applicant); // re-render
                        }
                    });
                }
                
                evalContainer.appendChild(panelEl);
            }

            // Render Consensus Panel if there are secondary reviewers
            if (Object.keys(secondaryReviewers).length > 0) {
                let conRatings = consensusRatings[currentApplicantEmail];
                let conNotes = consensusNotes[currentApplicantEmail];
                
                // If it hasn't been explicitly calculated or saved for this applicant, pre-populate it dynamically
                if (!conRatings) {
                    conRatings = {};
                    const counts = { bsc: 0, msc: 0, research: 0, prof: 0, english: 0, cv: 0 };
                    const sums = { bsc: 0, msc: 0, research: 0, prof: 0, english: 0, cv: 0 };
                    let combinedNotes = [];
                    
                    if (ratings[currentApplicantEmail]) {
                        Object.keys(sums).forEach(k => {
                            if (ratings[currentApplicantEmail][k] !== undefined) {
                                sums[k] += ratings[currentApplicantEmail][k];
                                counts[k]++;
                            }
                        });
                    }
                    if (reviewerNotes[currentApplicantEmail] && reviewerNotes[currentApplicantEmail].trim()) {
                        combinedNotes.push(`${primaryReviewerName || 'Primary'}:\n${reviewerNotes[currentApplicantEmail].trim()}`);
                    }
                    
                    Object.entries(secondaryReviewers).forEach(([name, data]) => {
                        if (data.ratings && data.ratings[currentApplicantEmail]) {
                            Object.keys(sums).forEach(k => {
                                if (data.ratings[currentApplicantEmail][k] !== undefined) {
                                    sums[k] += data.ratings[currentApplicantEmail][k];
                                    counts[k]++;
                                }
                            });
                        }
                        if (data.notes && data.notes[currentApplicantEmail] && data.notes[currentApplicantEmail].trim()) {
                            combinedNotes.push(`${name}:\n${data.notes[currentApplicantEmail].trim()}`);
                        }
                    });
                    
                    Object.keys(sums).forEach(k => {
                        if (counts[k] > 0) {
                            conRatings[k] = sums[k] / counts[k];
                        }
                    });
                    
                    conNotes = combinedNotes.join('\n\n');
                    
                    // Save dynamically calculated state so it persists
                    consensusRatings[currentApplicantEmail] = conRatings;
                    consensusNotes[currentApplicantEmail] = conNotes;
                    localStorage.setItem('consensusRatings', JSON.stringify(consensusRatings));
                    localStorage.setItem('consensusNotes', JSON.stringify(consensusNotes));
                    if (typeof updateApplicantList === 'function') updateApplicantList(); // refresh scores
                }
                
                // Calculate consensus score for display
                let conScore = 0;
                let totalWeight = 0;
                Object.keys(weights).forEach(k => {
                    totalWeight += weights[k];
                    conScore += (conRatings[k] || 0) * weights[k];
                });
                conScore = totalWeight > 0 ? (conScore / totalWeight) : 0;
                
                const consensusHtml = `
                    <section class="card evaluation-panel consensus-panel" style="flex: 1; min-width: 300px; margin-bottom: 0; border: 2px solid #3b82f6; background-color: #eff6ff;">
                        <div class="eval-header">
                            <div style="display: flex; align-items: center; justify-content: space-between; width: 100%;">
                                <h2>Consensus Evaluation</h2>
                            </div>
                            <div class="total-score">Score: <strong>${conScore.toFixed(2)}</strong> / 10</div>
                        </div>
                        <div class="eval-grid">
                            <div class="eval-input"><label>BSc Grade</label><input type="text" id="con_bsc" value="${conRatings.bsc !== undefined ? conRatings.bsc : ''}"></div>
                            <div class="eval-input"><label>MSc Grade</label><input type="text" id="con_msc" value="${conRatings.msc !== undefined ? conRatings.msc : ''}"></div>
                            <div class="eval-input"><label>Research Exp.</label><input type="text" id="con_research" value="${conRatings.research !== undefined ? conRatings.research : ''}"></div>
                            <div class="eval-input"><label>Prof. Exp.</label><input type="text" id="con_prof" value="${conRatings.prof !== undefined ? conRatings.prof : ''}"></div>
                            <div class="eval-input"><label>English Skills</label><input type="text" id="con_english" value="${conRatings.english !== undefined ? conRatings.english : ''}"></div>
                            <div class="eval-input"><label>CV & Cover Letter</label><input type="text" id="con_cv" value="${conRatings.cv !== undefined ? conRatings.cv : ''}"></div>
                        </div>
                        <div class="reviewer-notes-container mt-4">
                            <label class="reviewer-notes-label">Consensus Notes</label>
                            <textarea id="con_notes" class="reviewer-notes">${conNotes}</textarea>
                        </div>
                    </section>
                `;
                
                const template = document.createElement('template');
                template.innerHTML = consensusHtml.trim();
                const conPanelEl = template.content.firstChild;
                
                // Add event listeners to inputs to save to consensusRatings and consensusNotes
                const conInputs = {
                    bsc: conPanelEl.querySelector('#con_bsc'),
                    msc: conPanelEl.querySelector('#con_msc'),
                    research: conPanelEl.querySelector('#con_research'),
                    prof: conPanelEl.querySelector('#con_prof'),
                    english: conPanelEl.querySelector('#con_english'),
                    cv: conPanelEl.querySelector('#con_cv')
                };
                Object.keys(conInputs).forEach(k => {
                    conInputs[k].addEventListener('input', () => {
                        if (!currentApplicantEmail) return;
                        if (!consensusRatings[currentApplicantEmail]) consensusRatings[currentApplicantEmail] = {};
                        
                        let val = parseFloat(conInputs[k].value);
                        if (isNaN(val)) val = 0;
                        val = clampScore(val);
                        
                        consensusRatings[currentApplicantEmail][k] = val;
                        localStorage.setItem('consensusRatings', JSON.stringify(consensusRatings));
                        
                        // Recalculate score
                        let s = 0;
                        let tw = 0;
                        Object.keys(weights).forEach(wk => {
                            tw += weights[wk];
                            s += (consensusRatings[currentApplicantEmail][wk] || 0) * weights[wk];
                        });
                        const updatedScore = tw > 0 ? (s / tw) : 0;
                        conPanelEl.querySelector('.total-score strong').textContent = updatedScore.toFixed(2);
                    });
                });
                
                const conNotesEl = conPanelEl.querySelector('#con_notes');
                conNotesEl.addEventListener('input', () => {
                    if (!currentApplicantEmail) return;
                    consensusNotes[currentApplicantEmail] = conNotesEl.value;
                    localStorage.setItem('consensusNotes', JSON.stringify(consensusNotes));
                });
                
                // Add margins since it's below the container
                conPanelEl.style.marginTop = '20px';
                conPanelEl.style.width = '100%';
                evalContainer.insertAdjacentElement('afterend', conPanelEl);
            }
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
        const setUniversityWithRankingLink = (id, key, altKeys = [], countryKey = null, countryAltKeys = []) => {
            const el = document.getElementById(id);
            if (el) {
                const val = getVal(key, altKeys);
                if (val && val !== '/' && val.trim() !== '') {
                    const country = countryKey ? (getVal(countryKey, countryAltKeys) || '') : '';
                    const countryPart = (country && country !== '/' && country.trim()) ? ` in ${country.trim()}` : '';
                    const query = encodeURIComponent(`ranking${countryPart} timeshighereducation.com ${val}`);
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
        setUniversityWithRankingLink('detailBscUni', 'University (full name)', [], 'Education: Bachelor Country');
        setText('detailBscProgram', 'Study Program name');
        setText('detailBscGrade', 'Final Grade (in the format your_grade / maximum_grade)');

        // Education - MSc
        setText('detailMscCountry', 'Education: Master Country');
        setText('detailMscYear', 'Year');
        setUniversityWithRankingLink('detailMscUni', 'University (full name)_2', ['University (full name)2'], 'Education: Master Country');
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
                document.getElementById('detailThesisUrl').href = normalizeUrl(thesisUrl);
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
                files.forEach((fileName, idx) => {
                    if(!fileName) return;
                    const a = document.createElement('a');
                    a.className = 'doc-btn';
                    
                    // Try to find the file in the file map
                    let fileObject = null;
                    const possiblePaths = [
                        fileName,
                        (link ? link + '/' + fileName : null),
                        'data/' + (link ? link + '/' + fileName : fileName)
                    ];

                    for (let path of possiblePaths) {
                        if (path && fileMap.has(path)) {
                            fileObject = fileMap.get(path);
                            break;
                        }
                    }

                    if (fileObject) {
                        a.href = URL.createObjectURL(fileObject);
                    } else {
                        // Fallback to the old logic if fileMap is empty (e.g. on refresh)
                        const basePath = docBasePath.endsWith('/') ? docBasePath : docBasePath + '/';
                        if(link) {
                            a.href = basePath + link + '/' + fileName;
                        } else {
                            a.href = basePath + fileName;
                        }
                    }
                    a.target = '_blank';
                    
                    const displayLabel = files.length > 1 ? `${label} (${idx+1})` : label;
                    a.innerHTML = `${docIcon} <span title="${fileName}">${displayLabel}</span>`;
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
        
        if(extLink && extLink !== '/' && extLink.trim() !== '') {
            extDocsLink.href = normalizeUrl(extLink);
            extDocsContainer.classList.remove('hidden');
        } else {
            extDocsContainer.classList.add('hidden');
        }
    }

    const changeFolderBtn = document.getElementById("changeFolderBtn");

    function resetApp() {
        // Clear State
        localStorage.removeItem('cachedApplicants');
        localStorage.removeItem('cachedFilename');
        applicantsData = [];
        fileMap.clear();
        excelFiles = [];
        storedDirHandle = null;
        clearDirHandle();
        
        // Reset Landing Page UI
        if (folderUploadLabel) folderUploadLabel.textContent = "Select Survey Folder";
        if (discoveryPlaceholder) discoveryPlaceholder.classList.remove('hidden');
        discoveryPanel.classList.add('hidden');
        uploadStatus.textContent = "";
        folderUpload.value = ''; // Clear input

        // Hide reconnect banner if present
        const banner = document.getElementById('reconnectBanner');
        if (banner) banner.remove();

        // Reset Main App UI
        applicantList.innerHTML = '';
        applicantDetails.classList.add('hidden');
        welcomeMessage.classList.remove('hidden');
        landingPage.classList.remove("hidden");
        appContainer.classList.add("hidden");
        updateTopBarTitle("");
    }

    if (changeFolderBtn) changeFolderBtn.addEventListener('click', resetApp);

    // --- Close Document Logic ---
    const clearDataBtn = document.getElementById("clearDataBtn");
    if (clearDataBtn) {
        clearDataBtn.addEventListener('click', () => {
            resetApp();
            // Reset filters
            if (evalFilterSelect) evalFilterSelect.value = 'all';
            if (markedFilterSelect) markedFilterSelect.value = 'all';
        });
    }

    // --- Auto-load cached file & restore folder access ---
    const cachedFile = localStorage.getItem('cachedApplicants');
    if (cachedFile) {
        try {
            applicantsData = JSON.parse(cachedFile);
            originalFilename = localStorage.getItem('cachedFilename') || 'applicants';
            updateTopBarTitle(originalFilename);
            
            landingPage.classList.add("hidden");
            appContainer.classList.remove("hidden");
            updateApplicantList();

            // Try to restore the directory handle from IndexedDB
            restoreFolderAccess();
        } catch (e) {
            console.error("Error loading cached file data:", e);
        }
    }

    async function restoreFolderAccess() {
        try {
            const handle = await loadDirHandle();
            if (!handle) {
                showReconnectBanner('Document links are unavailable — no saved folder access.', true);
                return;
            }

            // Verify/request permission
            let perm = await handle.queryPermission({ mode: 'read' });
            if (perm === 'granted') {
                // Permission already granted – rebuild fileMap silently
                storedDirHandle = handle;
                fileMap.clear();
                await buildFileMapFromHandle(handle, '');
                console.log(`Restored folder access: ${fileMap.size} files mapped.`);
                return;
            }

            // Permission not yet granted – show banner so user can click to grant
            showReconnectBanner(
                `Document links need folder access to "${handle.name}".`,
                false,
                async () => {
                    perm = await handle.requestPermission({ mode: 'read' });
                    if (perm === 'granted') {
                        storedDirHandle = handle;
                        fileMap.clear();
                        await buildFileMapFromHandle(handle, '');
                        console.log(`Restored folder access: ${fileMap.size} files mapped.`);
                        const banner = document.getElementById('reconnectBanner');
                        if (banner) banner.remove();
                    }
                }
            );
        } catch (err) {
            console.warn('Could not restore folder access:', err);
            showReconnectBanner('Document links are unavailable — folder access could not be restored.', true);
        }
    }

    function showReconnectBanner(message, showSelectBtn, grantAction) {
        // Remove existing banner if any
        const existing = document.getElementById('reconnectBanner');
        if (existing) existing.remove();

        const banner = document.createElement('div');
        banner.id = 'reconnectBanner';
        banner.className = 'reconnect-banner';

        const icon = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>`;
        
        let buttonsHTML = '';
        if (grantAction) {
            buttonsHTML += `<button id="grantAccessBtn" class="banner-btn banner-btn-primary">Grant Access</button>`;
        }
        if (showSelectBtn) {
            buttonsHTML += `<button id="bannerSelectFolder" class="banner-btn banner-btn-secondary">Select Folder</button>`;
        }
        buttonsHTML += `<button id="dismissBanner" class="banner-btn banner-btn-dismiss" title="Dismiss">✕</button>`;

        banner.innerHTML = `${icon}<span>${message}</span><div class="banner-actions">${buttonsHTML}</div>`;

        // Insert banner right after the top bar
        const topBar = document.querySelector('.top-bar');
        if (topBar && topBar.nextSibling) {
            topBar.parentNode.insertBefore(banner, topBar.nextSibling);
        } else {
            appContainer.prepend(banner);
        }

        // Wire up buttons
        const grantBtn = document.getElementById('grantAccessBtn');
        if (grantBtn && grantAction) {
            grantBtn.addEventListener('click', grantAction);
        }

        const selectBtn = document.getElementById('bannerSelectFolder');
        if (selectBtn) {
            selectBtn.addEventListener('click', () => {
                if (window.showDirectoryPicker) {
                    (async () => {
                        await selectFolderViaFSAPI();
                        if (fileMap.size > 0) {
                            banner.remove();
                        }
                    })();
                } else {
                    // Trigger the file input
                    folderUpload.click();
                    folderUpload.addEventListener('change', function handler() {
                        if (fileMap.size > 0) banner.remove();
                        folderUpload.removeEventListener('change', handler);
                    });
                }
            });
        }

        const dismissBtn = document.getElementById('dismissBanner');
        if (dismissBtn) {
            dismissBtn.addEventListener('click', () => banner.remove());
        }
    }
});
