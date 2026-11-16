// Quiz v21.1 - Local-First Cloud-Sync + UI Overlap Fix + Resume Quiz on Refresh
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { getAuth, setPersistence, browserLocalPersistence,
         signInWithEmailAndPassword, signOut, onAuthStateChanged }
       from "https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js";
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager,
         doc, getDoc, setDoc, updateDoc, collection, getDocs, query, orderBy, limit }
       from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

// ⚠️ உங்கள் FIREBASE CONFIG
const firebaseConfig = {
    apiKey: "AIzaSyCzM21SGVy2DZm7c0X8EB0CJcnhZ7mNBn8",
    authDomain: "webapp-cd3a5.firebaseapp.com",
    projectId: "webapp-cd3a5",
    storageBucket: "webapp-cd3a5.firebasestorage.app",
    messagingSenderId: "YOUR_SENDER_ID",
    appId: "YOUR_APP_ID"
};

const app  = initializeApp(firebaseConfig);
const auth = getAuth(app);
await setPersistence(auth, browserLocalPersistence);
const db   = initializeFirestore(app, {
    localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() })
});

// DOM Elements
const questionElement = document.getElementById('question');
const optionElements = document.querySelectorAll('.option');
const feedbackElement = document.getElementById('feedback');
const prevButton = document.getElementById('prev-btn');
const nextButton = document.getElementById('next-btn');
const scoreElement = document.getElementById('score');
const totalQuestionsElement = document.getElementById('total-questions');
const progressBar = document.getElementById('progress-bar');
const endQuizElement = document.getElementById('end-quiz');
const finalScoreElement = document.getElementById('final-score');
const finalTotalElement = document.getElementById('final-total');
const restartButton = document.getElementById('restart-btn');
const exitButton = document.getElementById('exit-btn');
const timerContainer = document.getElementById('timer-container');
const timerElement = document.getElementById('timer');
const quizContent = document.getElementById('quiz-content');
const scoreSummaryElement = document.getElementById('score-summary');
const imageElement = document.getElementById('questionImage');

const filterContainer = document.getElementById('filter-container');
const yearButtonsContainer = document.getElementById('year-buttons');
const subjectButtonsContainer = document.getElementById('subject-buttons');
const typeButtonsContainer = document.getElementById('type-buttons');
const randomizeCheckbox = document.getElementById('randomize-checkbox');
const applyFiltersBtn = document.getElementById('apply-filters-btn');
const resetFiltersBtn = document.getElementById('reset-filters-btn');
const injectWrongCheckbox = document.getElementById('inject-wrong-checkbox');

const loginSection = document.getElementById('login-section');
const userInfoSection = document.getElementById('user-info-section');
const lockScreen = document.getElementById('lock-screen');
const mainScoreBoard = document.getElementById('main-score-board');

// State
let currentUid = null;
let allQuestions = [];
let wrongQuestions = {};
let quizData = [];
let currentQuestionIndex = 0;
let score = 0;
let userAnswers = {};
let questionOrder = [];
let subjectWiseScores = {};
let timerInterval = null;
let timeLeft = 30;
let questionStartTime = 0;
let sessionStartTime = 0;
let isRepairMode = false;

// ✅ சாதனத்தின் விவரங்களைச் சேமிக்க
let currentDeviceInfo = { autoName: "", customName: "" };

let availableYears = [], availableSubjects = [], availableTypes = [];
let selectedYears = [], selectedSubjects = [], selectedTypes = [];

// STAGE CONFIG
const STAGE_CONFIG = {
    1: { maxStars: 3,  mode: 'optional'   },
    2: { maxStars: 5, mode: 'compulsory' },
    3: { maxStars: 1, mode: 'locked'     },
    4: { maxStars: 2, mode: 'locked'     },
    5: { maxStars: 3, mode: 'locked'     }
};

// 🔥 UI Manager (எல்லா திரைகளையும் கட்டுப்படுத்த)
function hideAllScreens() {
    loginSection.classList.add('hidden');
    userInfoSection.classList.add('hidden');
    lockScreen.classList.add('hidden');
    mainScoreBoard.style.display = 'none';
    filterContainer.style.display = 'none';
    quizContent.style.display = 'none';
    endQuizElement.style.display = 'none';
}

function showLoginScreen() {
    hideAllScreens();
    loginSection.classList.remove('hidden');
}

function showFilterScreen() {
    hideAllScreens();
    userInfoSection.classList.remove('hidden');
    mainScoreBoard.style.display = 'block';
    filterContainer.style.display = 'block';
}

function showLockScreen() {
    hideAllScreens();
    userInfoSection.classList.remove('hidden');
    lockScreen.classList.remove('hidden');
    document.getElementById('locked-count').textContent = Object.keys(wrongQuestions).filter(id => wrongQuestions[id].weightage >= 3 && !isMastered(wrongQuestions[id])).length;
}

// 🔥 Active Quiz State (Refresh ஆனாலும் தேர்வு போகாமல் இருக்க)
function saveActiveQuizState() {
    if (!currentUid || quizData.length === 0) return;
    const state = {
        quizData, currentQuestionIndex, score, userAnswers, 
        sessionStartTime, isRepairMode, questionOrder, subjectWiseScores
    };
    localStorage.setItem(`activeQuiz_${currentUid}`, JSON.stringify(state));
}

function clearActiveQuizState() {
    if (currentUid) localStorage.removeItem(`activeQuiz_${currentUid}`);
}

function checkAndRestoreQuiz() {
    const saved = localStorage.getItem(`activeQuiz_${currentUid}`);
    if (saved) {
        const state = JSON.parse(saved);
        // ✅ இங்கே தான் confirm() கேட்கிறது
        if (confirm("உங்கள் முந்தைய தேர்வு முடிக்கப்படவில்லை. அதைத் தொடரவா?")) {
            quizData = state.quizData;
            currentQuestionIndex = state.currentQuestionIndex;
            score = state.score;
            userAnswers = state.userAnswers;
            sessionStartTime = state.sessionStartTime;
            isRepairMode = state.isRepairMode;
            questionOrder = state.questionOrder;
            subjectWiseScores = state.subjectWiseScores || {};
            
            hideAllScreens();
            userInfoSection.classList.remove('hidden');
            quizContent.style.display = 'block'; 
            timerContainer.style.display = 'block';
            questionElement.style.display = 'block';
            document.querySelector('.options').style.display = 'grid';
            feedbackElement.style.display = 'block'; 
            prevButton.style.display = 'block';
            nextButton.style.display = 'block';
            document.querySelector('.progress').style.display = 'block';
            document.querySelector('.score-container').style.display = 'block';
            
            totalQuestionsElement.textContent = quizData.length;
            scoreElement.textContent = score;
            displayQuestion();
            return true;
        } else {
            // 'Cancel' அழுத்தினால் தேர்வு நிலை அழிக்கப்படும்
            clearActiveQuizState();
        }
    }
    return false;
}

// 🔥 Device Detect & Register Functions
async function detectDevice() {
    let deviceStr = "Unknown Device";
    try {
        if (navigator.userAgentData && navigator.userAgentData.getHighEntropyValues) {
            const uaData = await navigator.userAgentData.getHighEntropyValues(["model", "platform", "platformVersion"]);
            const platform = uaData.platform || "Unknown OS";
            const model = uaData.model || "";
            const platformVer = uaData.platformVersion || "";
            
            if (model) {
                deviceStr = `${platform} (${model})`; 
            } else if (platform === "Windows") {
                const majorVer = parseInt(platformVer.split('.')[0], 10);
                deviceStr = (majorVer >= 13) ? "Windows 11" : "Windows 10";
            } else if (platform === "macOS") {
                deviceStr = "Mac"; 
            } else {
                deviceStr = `${platform} ${platformVer.split('.')[0]}`;
            }
        } else {
            const ua = navigator.userAgent;
            if (/iPhone|iPad/.test(ua)) deviceStr = "Apple iOS Device";
            else if (/Android/.test(ua)) {
                const match = ua.match(/Android\s([0-9.]+)/);
                deviceStr = `Android ${match ? match[1] : ""}`;
            } else if (/Windows/.test(ua)) deviceStr = "Windows PC";
            else if (/Mac/.test(ua)) deviceStr = "Mac";
            else if (/Linux/.test(ua)) deviceStr = "Linux";
        }
    } catch (e) { console.warn(e); }

    const ua = navigator.userAgent;
    const browser = /Edg/i.test(ua) ? "Edge" : /Chrome/i.test(ua) ? "Chrome" : /Safari/i.test(ua) ? "Safari" : /Firefox/i.test(ua) ? "Firefox" : "Browser";
    return `${deviceStr} - ${browser}`;
}

// 🔥 1. Custom Dialog UI Controller
function showDeviceDialog(autoName, devices) {
    return new Promise((resolve) => {
        const dialog = document.getElementById('device-dialog');
        const step1 = document.getElementById('device-step-1');
        const step2 = document.getElementById('device-step-2');
        const step3 = document.getElementById('device-step-3');
        const title = document.getElementById('device-dialog-title');
        const selectedDisplay = document.getElementById('dialog-selected-display');
        const selectedNameSpan = document.getElementById('selected-device-name');
        
        document.getElementById('dialog-auto-name').textContent = autoName;
        document.getElementById('dialog-device-input').value = "";
        document.getElementById('dialog-device-list').innerHTML = "";
        selectedDisplay.style.display = 'none';
        
        let selectedDevice = null;

        // Reset steps
        step1.style.display = 'block';
        step2.style.display = 'none';
        step3.style.display = 'none';
        title.textContent = "புதிய சாதனமா?";
        dialog.style.display = 'flex';

        // Cleanup old listeners
        const btnYes = document.getElementById('dialog-btn-yes');
        const btnNo = document.getElementById('dialog-btn-no');
        const btnSave = document.getElementById('dialog-btn-save');
        const btnBack1 = document.getElementById('dialog-btn-back-1');
        const btnBack2 = document.getElementById('dialog-btn-back-2');
        const btnConfirm = document.getElementById('dialog-btn-confirm');
        
        btnYes.replaceWith(btnYes.cloneNode(true));
        btnNo.replaceWith(btnNo.cloneNode(true));
        btnSave.replaceWith(btnSave.cloneNode(true));
        btnBack1.replaceWith(btnBack1.cloneNode(true));
        btnBack2.replaceWith(btnBack2.cloneNode(true));
        btnConfirm.replaceWith(btnConfirm.cloneNode(true));

        const newConfirmBtn = document.getElementById('dialog-btn-confirm');

        // 1. Yes Button
        document.getElementById('dialog-btn-yes').addEventListener('click', () => {
            step1.style.display = 'none';
            step2.style.display = 'block';
            title.textContent = "புதிய சாதனத்தின் பெயர்:";
            document.getElementById('dialog-device-input').focus();
        });

        // 2. No Button (Show List)
        document.getElementById('dialog-btn-no').addEventListener('click', () => {
            step1.style.display = 'none';
            step3.style.display = 'block';
            title.textContent = "உங்கள் சாதனத்தைத் தேர்ந்தெடுக்கவும்:";
            
            const listDiv = document.getElementById('dialog-device-list');
            
            // ✅ புதிய பிழை திருத்தம்: பட்டியலை முதலில் காலி செய் (Duplicate-ஐ தடுக்க)
            listDiv.innerHTML = ""; 
            
            devices.forEach((d) => {
                const btn = document.createElement('button');
                btn.className = 'filter-btn';
                btn.style.width = '100%';
                btn.style.marginBottom = '5px';
                btn.textContent = `${d.customName} (${d.autoName})`;
                
                btn.onclick = () => {
                    listDiv.querySelectorAll('button').forEach(b => b.classList.remove('selected'));
                    btn.classList.add('selected');
                    
                    selectedDevice = d;
                    selectedDisplay.style.display = 'block';
                    selectedNameSpan.textContent = `${d.customName} (${d.autoName})`;
                    newConfirmBtn.disabled = false;
                    newConfirmBtn.style.backgroundColor = '#40c057';
                    newConfirmBtn.style.cursor = 'pointer';
                };
                listDiv.appendChild(btn);
            });
        });

        // 3. Save New Device (Confirm)
        document.getElementById('dialog-btn-save').addEventListener('click', () => {
            const customName = document.getElementById('dialog-device-input').value.trim();
            if (!customName) { alert("தயவுசெய்து ஒரு பெயரை உள்ளிடவும்"); return; }
            dialog.style.display = 'none';
            resolve({ isNew: true, customName: customName, autoName: autoName });
        });

        // 4. Confirm Existing Device
        newConfirmBtn.addEventListener('click', () => {
            if (!selectedDevice) return;
            dialog.style.display = 'none';
            resolve({ isNew: false, device: selectedDevice });
        });

        // 5. Back Buttons
        document.getElementById('dialog-btn-back-1').addEventListener('click', () => {
            step2.style.display = 'none';
            step1.style.display = 'block';
            title.textContent = "புதிய சாதனமா?";
        });

        document.getElementById('dialog-btn-back-2').addEventListener('click', () => {
            step3.style.display = 'none';
            step1.style.display = 'block';
            title.textContent = "புதிய சாதனமா?";
            selectedDevice = null;
            newConfirmBtn.disabled = true;
            newConfirmBtn.style.backgroundColor = '#ced4da';
            newConfirmBtn.style.cursor = 'not-allowed';
            selectedDisplay.style.display = 'none';
        });
    });
}

// 🔥 2. Device Register Logic (Using the Dialog)
async function registerDevice(uid) {
    const localData = localStorage.getItem(`deviceInfo_${uid}`);
    const autoName = await detectDevice();
    const deviceId = autoName.replace(/[^a-zA-Z0-9]/g, '_'); 
    
    try {
        const deviceRef = doc(db, 'users', uid, 'devices', deviceId);
        const docSnap = await getDoc(deviceRef);
        const devicesColl = collection(db, 'users', uid, 'devices');
        const allDevicesSnap = await getDocs(devicesColl);

        if (localData) {
            currentDeviceInfo = JSON.parse(localData);
            currentDeviceInfo.lastUsed = new Date().toISOString();
            localStorage.setItem(`deviceInfo_${uid}`, JSON.stringify(currentDeviceInfo));
            setDoc(deviceRef, { lastUsed: currentDeviceInfo.lastUsed }, { merge: true }).catch(e => console.warn(e));
            return;
        }

        if (!allDevicesSnap.empty) {
            // OS Filter
            const currentOS = autoName.split(' - ')[0]; 
            let deviceList = [];
            let seenNames = new Set(); 

            allDevicesSnap.forEach(d => {
                const data = d.data();
                if (data.autoName && data.autoName.includes(currentOS)) {
                    const uniqueName = `${data.customName} (${data.autoName})`;
                    if (!seenNames.has(uniqueName)) {
                        seenNames.add(uniqueName);
                        deviceList.push({ id: d.id, ...data });
                    }
                }
            });

            if (deviceList.length > 0) {
                // Show Custom Dialog and Wait for User Choice
                const result = await showDeviceDialog(autoName, deviceList);
                
                if (result.isNew) {
                    const newDeviceId = `${deviceId}_${Date.now()}`; 
                    currentDeviceInfo = { autoName: result.autoName, customName: result.customName, lastUsed: new Date().toISOString() };
                    localStorage.setItem(`deviceInfo_${uid}`, JSON.stringify(currentDeviceInfo));
                    await setDoc(doc(db, 'users', uid, 'devices', newDeviceId), currentDeviceInfo);
                } else {
                    currentDeviceInfo = result.device;
                    currentDeviceInfo.lastUsed = new Date().toISOString();
                    localStorage.setItem(`deviceInfo_${uid}`, JSON.stringify(currentDeviceInfo));
                    await setDoc(doc(db, 'users', uid, 'devices', currentDeviceInfo.id), { lastUsed: currentDeviceInfo.lastUsed }, { merge: true });
                }
                return;
            }
        }

        // Local காலி, Firebase-லும் (அந்த OS-க்கு) இல்லை
        const result = await showDeviceDialog(autoName, []);
        const newDeviceId = `${deviceId}_${Date.now()}`; 
        currentDeviceInfo = { autoName: result.autoName, customName: result.customName, lastUsed: new Date().toISOString() };
        localStorage.setItem(`deviceInfo_${uid}`, JSON.stringify(currentDeviceInfo));
        await setDoc(doc(db, 'users', uid, 'devices', newDeviceId), currentDeviceInfo);

    } catch (e) {
        console.error("Device error:", e);
        // Fallback to offline mode
        if (localData) {
            currentDeviceInfo = JSON.parse(localData);
        } else {
            const customName = autoName;
            currentDeviceInfo = { autoName: autoName, customName: customName, lastUsed: new Date().toISOString() };
            localStorage.setItem(`deviceInfo_${uid}`, JSON.stringify(currentDeviceInfo));
        }
    }
}

// 🔥 Star System Logic
function isMastered(wq) {
    if (!wq) return false;
    const cfg = STAGE_CONFIG[wq.weightage] || STAGE_CONFIG[5];
    return wq.stars >= cfg.maxStars; 
}

function applyAnswer(wq, isCorrect) {
    if (!wq) return isCorrect ? null : { weightage:1, stars:0, wrongCount:1, correctCount:0 };
    
    const cfg = STAGE_CONFIG[wq.weightage] || STAGE_CONFIG[5];
    const atMax = wq.stars >= cfg.maxStars;

    if (isCorrect) {
        return { ...wq, stars: Math.min(wq.stars + 1, cfg.maxStars), correctCount: (wq.correctCount || 0) + 1 };
    } else {
        if (atMax) {
            if (wq.weightage < 5) {
                return { ...wq, weightage: wq.weightage + 1, stars: 0, wrongCount: (wq.wrongCount || 0) + 1 };
            } else {
                return { ...wq, stars: 0, wrongCount: (wq.wrongCount || 0) + 1 };
            }
        } else {
            return { ...wq, stars: Math.max(wq.stars - 1, 0), wrongCount: (wq.wrongCount || 0) + 1 };
        }
    }
}

function isLocked(wqMap) {
    return Object.values(wqMap).some(q => q.weightage >= 3 && !isMastered(q));
}

// 🔥 Question Bank Load (Local First + Version Check)
async function loadQuestionBank() {
    const cachedStr = localStorage.getItem('qbCache');
    const cachedVer = localStorage.getItem('qbVersion');
    let cached = null;
    if (cachedStr) cached = JSON.parse(cachedStr);

    try {
        const vRes = await fetch(`./data/question_bank_version.json?t=${Date.now()}`);
        const { version } = await vRes.json();
        
        if (version !== cachedVer || !cached) {
            console.log("New version found. Downloading Excel...");
            const xRes = await fetch(`./data/question_bank.xlsx?t=${Date.now()}`);
            const arrayBuffer = await xRes.arrayBuffer();
            const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
            const sheetName = workbook.SheetNames.includes('MASTER_DATA') ? 'MASTER_DATA' : workbook.SheetNames[0];
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
            
            localStorage.setItem('qbCache', JSON.stringify(jsonData));
            localStorage.setItem('qbVersion', version);
            processExcelData(jsonData);
        } else {
            console.log("Using Cached Question Bank.");
            processExcelData(cached);
        }
    } catch (error) {
        console.warn('Offline mode: Loading from cache', error);
        if (cached) processExcelData(cached);
        else alert('வினா வங்கியை ஏற்ற முடியவில்லை. இணையத்தை சரிபார்க்கவும்.');
    }
}

// 🔥 Process Excel
function processExcelData(data) {
    const processText = (text) => {
        if (!text) return "";
        return String(text).replace(/\n/g, '<br>');
    };
    const extractYear = (idString) => {
        if (!idString) return "வருடம்";
        const parts = String(idString).trim().split('_');
        if (parts.length > 1) {
            parts.pop(); 
            return parts.join('_').toUpperCase(); 
        }
        return "வருடம்";
    };

    allQuestions = data.map(row => {
        return {
            id: String(row.id || '').trim(),
            year: extractYear(row.id), 
            subject: processText(row.subject) || "பாடம்",
            type: processText(row.type) || "வடிவம்",
            question: processText(row.question) || "கேள்வி இல்லை",
            image_path: processText(row.image_path) || "",
            options: {
                A: processText(row.a) || "விருப்பம் A",
                B: processText(row.b) || "விருப்பம் B",
                C: processText(row.c) || "விருப்பம் C",
                D: processText(row.d) || "விருப்பம் D"
            },
            key: String(row.key || "A").trim().toUpperCase()
        };
    }).filter(item => item.question !== "கேள்வி இல்லை");

    if (allQuestions.length > 0) {
        extractFilters(allQuestions);
    }
}

// 🔥 Filters
function extractFilters(questions) {
    availableYears = ['all']; availableSubjects = ['all']; availableTypes = ['all'];
    questions.forEach(item => {
        if (item.year && !availableYears.includes(item.year)) availableYears.push(item.year);
        if (item.subject && !availableSubjects.includes(item.subject)) availableSubjects.push(item.subject);
        if (item.type && !availableTypes.includes(item.type)) availableTypes.push(item.type);
    });
    
    yearButtonsContainer.innerHTML = '';
    availableYears.sort().forEach(year => {
        const btn = document.createElement('button');
        btn.className = 'filter-btn'; btn.textContent = year === 'all' ? 'அனைத்தும்' : year;
        btn.setAttribute('data-value', year);
        btn.addEventListener('click', () => toggleFilter('year', year, btn));
        yearButtonsContainer.appendChild(btn);
    });
    
    subjectButtonsContainer.innerHTML = '';
    availableSubjects.sort().forEach(sub => {
        const btn = document.createElement('button');
        btn.className = 'filter-btn'; btn.textContent = sub === 'all' ? 'அனைத்தும்' : sub;
        btn.setAttribute('data-value', sub);
        btn.addEventListener('click', () => toggleFilter('subject', sub, btn));
        subjectButtonsContainer.appendChild(btn);
    });
    
    typeButtonsContainer.innerHTML = '';
    availableTypes.sort().forEach(type => {
        const btn = document.createElement('button');
        btn.className = 'filter-btn'; btn.textContent = type === 'all' ? 'அனைத்தும்' : type;
        btn.setAttribute('data-value', type);
        btn.addEventListener('click', () => toggleFilter('type', type, btn));
        typeButtonsContainer.appendChild(btn);
    });
    
    loadSavedFilters();
}

function toggleFilter(type, value, button) {
    if (value === 'all') {
        if (type === 'year') {
            selectedYears = [...availableYears];
            document.querySelectorAll('#year-buttons .filter-btn').forEach(btn => btn.classList.add('selected'));
        } else if (type === 'subject') {
            selectedSubjects = [...availableSubjects];
            document.querySelectorAll('#subject-buttons .filter-btn').forEach(btn => btn.classList.add('selected'));
        } else {
            selectedTypes = [...availableTypes];
            document.querySelectorAll('#type-buttons .filter-btn').forEach(btn => btn.classList.add('selected'));
        }
        return;
    }
    button.classList.toggle('selected');
    if (type === 'year') {
        if (button.classList.contains('selected')) selectedYears.push(value); else selectedYears = selectedYears.filter(y => y !== value);
        const allBtn = document.querySelector('#year-buttons .filter-btn[data-value="all"]');
        if (selectedYears.length === availableYears.length - 1) { allBtn.classList.add('selected'); selectedYears.push('all'); } 
        else { allBtn.classList.remove('selected'); selectedYears = selectedYears.filter(y => y !== 'all'); }
    } else if (type === 'subject') {
        if (button.classList.contains('selected')) selectedSubjects.push(value); else selectedSubjects = selectedSubjects.filter(s => s !== value);
        const allBtn = document.querySelector('#subject-buttons .filter-btn[data-value="all"]');
        if (selectedSubjects.length === availableSubjects.length - 1) { allBtn.classList.add('selected'); selectedSubjects.push('all'); } 
        else { allBtn.classList.remove('selected'); selectedSubjects = selectedSubjects.filter(s => s !== 'all'); }
    } else {
        if (button.classList.contains('selected')) selectedTypes.push(value); else selectedTypes = selectedTypes.filter(t => t !== value);
        const allBtn = document.querySelector('#type-buttons .filter-btn[data-value="all"]');
        if (selectedTypes.length === availableTypes.length - 1) { allBtn.classList.add('selected'); selectedTypes.push('all'); } 
        else { allBtn.classList.remove('selected'); selectedTypes = selectedTypes.filter(t => t !== 'all'); }
    }
}

function loadSavedFilters() {
    const saved = localStorage.getItem('quizFilters');
    if (saved) {
        const f = JSON.parse(saved);
        if (f.years) { selectedYears = f.years; document.querySelectorAll('#year-buttons .filter-btn').forEach(btn => { if (selectedYears.includes(btn.getAttribute('data-value'))) btn.classList.add('selected'); }); }
        if (f.subjects) { selectedSubjects = f.subjects; document.querySelectorAll('#subject-buttons .filter-btn').forEach(btn => { if (selectedSubjects.includes(btn.getAttribute('data-value'))) btn.classList.add('selected'); }); }
        if (f.types) { selectedTypes = f.types; document.querySelectorAll('#type-buttons .filter-btn').forEach(btn => { if (selectedTypes.includes(btn.getAttribute('data-value'))) btn.classList.add('selected'); }); }
        if (f.randomize !== undefined) randomizeCheckbox.checked = f.randomize;
    } else {
        selectedYears = [...availableYears]; selectedSubjects = [...availableSubjects]; selectedTypes = [...availableTypes];
        document.querySelectorAll('.filter-btn').forEach(btn => btn.classList.add('selected'));
    }
}

// 🔥 Options-ஐ கலைக்கும் (Shuffle) புதிய ஃபங்ஷன்
function shuffleOptions(questions) {
    return questions.map(q => {
        if (!q.options || !q.key) return q;
        
        // ஒவ்வொரு விருப்பத்தையும் அதன் சரியான விடையுடன் இணைக்கவும்
        let optObjs = [
            { text: q.options.A, isCorrect: q.key === 'A' },
            { text: q.options.B, isCorrect: q.key === 'B' },
            { text: q.options.C, isCorrect: q.key === 'C' },
            { text: q.options.D, isCorrect: q.key === 'D' }
        ];
        
        // விருப்பங்களைக் கலைக்கவும் (Randomize)
        optObjs.sort(() => Math.random() - 0.5);
        
        let newOptions = {};
        let newKey = 'A';
        
        // புதிய வரிசைப்படி A, B, C, D என மீண்டும் ஒதுக்கவும்
        optObjs.forEach((obj, i) => {
            const letter = String.fromCharCode(65 + i); // 65 என்பது 'A'
            newOptions[letter] = obj.text;
            if (obj.isCorrect) newKey = letter; // புதிய சரியான விடை எழுத்து
        });
        
        return { ...q, options: newOptions, key: newKey };
    });
}

// 🔥 Apply Filters
function applyFilters() {
    const randomize = randomizeCheckbox.checked;
    const inject = injectWrongCheckbox ? injectWrongCheckbox.checked : false;
    localStorage.setItem('quizFilters', JSON.stringify({ years: selectedYears, subjects: selectedSubjects, types: selectedTypes, randomize }));

    let filtered = allQuestions;
    if (selectedYears.length > 0 && !selectedYears.includes('all')) filtered = filtered.filter(q => selectedYears.includes(q.year));
    if (selectedSubjects.length > 0 && !selectedSubjects.includes('all')) filtered = filtered.filter(q => selectedSubjects.includes(q.subject));
    if (selectedTypes.length > 0 && !selectedTypes.includes('all')) filtered = filtered.filter(q => selectedTypes.includes(q.type));
    
    const wrongIds = new Set();
    Object.entries(wrongQuestions).forEach(([id, wq]) => {
        if (isMastered(wq)) return;
        if (wq.weightage >= 2) wrongIds.add(id);
        else if (wq.weightage === 1 && inject) wrongIds.add(id);
    });

    const injected = allQuestions.filter(q => {
        if (!wrongIds.has(q.id)) return false;
        const wq = wrongQuestions[q.id];
        if (wq.weightage >= 2) return true;
        if (!selectedYears.includes('all') && !selectedYears.includes(q.year)) return false;
        if (!selectedSubjects.includes('all') && !selectedSubjects.includes(q.subject)) return false;
        if (!selectedTypes.includes('all') && !selectedTypes.includes(q.type)) return false;
        return true;
    });

    const seen = new Set(filtered.map(q => q.id));
    injected.forEach(q => { if (!seen.has(q.id)) { filtered.push(q); seen.add(q.id); } });

    if (filtered.length === 0) { alert('கேள்விகள் இல்லை'); return; }
    
    quizData = filtered;
    if (randomize) {
        quizData = quizData.sort(() => Math.random() - 0.5); // கேள்விகளைக் கலை
        quizData = shuffleOptions(quizData); // ✅ விருப்பங்களையும் கலை
    }
    
    hideAllScreens();
    userInfoSection.classList.remove('hidden');
    startQuiz();
}

function resetFilters() {
    selectedYears = [...availableYears]; selectedSubjects = [...availableSubjects]; selectedTypes = [...availableTypes];
    document.querySelectorAll('.filter-btn').forEach(btn => btn.classList.add('selected'));
    randomizeCheckbox.checked = true; if(injectWrongCheckbox) injectWrongCheckbox.checked = false;
}

// 🔥 LocalStorage & Firebase Sync Functions
function saveWQLocal(uid, wq) { localStorage.setItem(`wq_${uid}`, JSON.stringify(wq)); }
function loadWQLocal(uid) { return JSON.parse(localStorage.getItem(`wq_${uid}`) || '{}'); }

function saveSessionLocal(uid, session) {
    let sessions = loadSessionsLocal(uid);
    sessions.push(session);
    if (sessions.length > 500) sessions = sessions.slice(-500);
    localStorage.setItem(`sessions_${uid}`, JSON.stringify(sessions));
}
function loadSessionsLocal(uid) { return JSON.parse(localStorage.getItem(`sessions_${uid}`) || '[]'); }

async function syncWQToCloud(uid, qid, data) {
    try { await setDoc(doc(db, 'users', uid, 'wrongQuestions', qid), data, { merge: true }); } 
    catch (e) { console.warn("WQ Cloud Sync Failed (Offline)", e); }
}
async function syncSessionToCloud(uid, sid, data) {
    try { await setDoc(doc(db, 'users', uid, 'quizSessions', sid), data); } 
    catch (e) { console.warn("Session Cloud Sync Failed (Offline)", e); }
}

// 🔥 Auth & Sync on Login
document.getElementById('login-btn').addEventListener('click', async () => {
    const email = document.getElementById('login-email').value.trim();
    const pass  = document.getElementById('login-password').value;
    const errEl = document.getElementById('auth-error');
    errEl.style.display = 'none';
    try { await signInWithEmailAndPassword(auth, email, pass); } 
    catch (e) { errEl.textContent = 'உள்நுழைவு தோல்வி: ' + e.message; errEl.style.display = 'block'; }
});

// ✅ Logout செய்யும் போது தேர்வு நிலை (State) அழிக்கப்படாமல் சேமிக்கப்படும்
document.getElementById('logout-btn').addEventListener('click', () => {
    // clearActiveQuizState(); // ❌ இதை நீக்கிவிட்டோம், அதனால் தேர்வு நிலை அழியாது
    signOut(auth);
});

onAuthStateChanged(auth, async user => {
    if (!user) {
        showLoginScreen();
        return;
    }
    
    // லாகின் ஆனதும் மற்ற திரைகளை மறை
    hideAllScreens(); 
    userInfoSection.classList.remove('hidden');
    currentUid = user.uid;
    
    const emailDisplay = document.getElementById('user-email-display');
    if (emailDisplay) emailDisplay.textContent = user.email;
    
    // Local Data Load
    wrongQuestions = loadWQLocal(currentUid);
    renderDashboard(loadSessionsLocal(currentUid)); 
    await registerDevice(currentUid);
    await loadQuestionBank(); // Excel லோட் செய்

    // Cloud Sync
    try {
        const snapWQ = await getDocs(collection(db, 'users', currentUid, 'wrongQuestions'));
        let cloudWQ = {}; snapWQ.forEach(d => { cloudWQ[d.id] = d.data(); });
        wrongQuestions = { ...wrongQuestions, ...cloudWQ };
        saveWQLocal(currentUid, wrongQuestions);

        const qSess = query(collection(db, 'users', currentUid, 'quizSessions'), orderBy('startTime', 'desc'), limit(100));
        const snapSess = await getDocs(qSess);
        let cloudSessions = []; snapSess.forEach(d => cloudSessions.push(d.data()));
        
        if (cloudSessions.length > 0) {
            const allSessions = [...loadSessionsLocal(currentUid), ...cloudSessions];
            const uniqueSessions = Array.from(new Map(allSessions.map(s => [s.sessionId, s])).values());
            uniqueSessions.sort((a,b) => new Date(b.startTime) - new Date(a.startTime));
            
            localStorage.setItem(`sessions_${currentUid}`, JSON.stringify(uniqueSessions));
            renderDashboard(uniqueSessions);
        }
    } catch (e) {
        console.warn("Offline Mode: Using Local Data only.", e);
    }

    // ✅ திரை கட்டுப்பாடு & Refresh பிழை திருத்தம்
    // 1. பழைய தேர்வு இருந்தால் Restore செய்
    const isRestored = checkAndRestoreQuiz();

    // 2. Restore ஆகவில்லை என்றால், லாக் உள்ளதா பார், இல்லைனா ஃபில்டர் காட்டு
    if (!isRestored) {
        if (isLocked(wrongQuestions)) {
            showLockScreen();
        } else {
            showFilterScreen();
        }
    }
});

// 🔥 Render Dashboard (Local Data Based)
function renderDashboard(sessions) {
    if (!sessions.length) {
        document.getElementById('score-details').innerHTML = '<p>இதுவரை மதிப்பெண்கள் இல்லை</p>';
        document.getElementById('average-score').textContent = '0';
        document.getElementById('main-score-board').style.display = 'block';
        return;
    }
    
    let totalScore = 0, totalQ = 0, subj = {};
    const todayStr = new Date().toISOString().slice(0, 10);
    let todayCount = 0;

    sessions.forEach(s => {
        totalScore += s.correctCount || 0; totalQ += s.totalQuestions || 0;
        if (s.date === todayStr) todayCount++; 
        
        (s.subjects || []).forEach(sub => {
            if (!subj[sub.name]) subj[sub.name] = { s: 0, t: 0 };
            subj[sub.name].s += sub.score; subj[sub.name].t += sub.total;
        });
    });
    
    const overallAvg = totalQ > 0 ? Math.round((totalScore / totalQ) * 100) : 0;
    document.getElementById('average-score').textContent = overallAvg;
    
    let html = `<div class="subject-score"><span><b>இன்றைய தேர்வுகள்:</b></span><span><b>${todayCount}</b></span></div><hr style="margin: 10px 0; border: 0; border-top: 1px solid #ced4da;">`;
    
    for (const [name, data] of Object.entries(subj)) {
        const avg = data.t > 0 ? Math.round((data.s / data.t) * 100) : 0;
        html += `<div class="subject-score"><span>${name}:</span><span>${avg}% (${data.s}/${data.t})</span></div>`;
    }
    html += `<div class="subject-score"><span><b>மொத்த தேர்வுகள்: ${sessions.length}</b></span><span><b>${overallAvg}% (${totalScore}/${totalQ})</b></span></div>`;
    
    document.getElementById('score-details').innerHTML = html;
    document.getElementById('main-score-board').style.display = 'block';
}

// 🔥 Quiz Logic
function startQuiz() {
    clearActiveQuizState();
    subjectWiseScores = {};
    questionOrder = Array.from({length: quizData.length}, (_, i) => i);
    currentQuestionIndex = 0; score = 0; userAnswers = {}; sessionStartTime = Date.now();
    scoreElement.textContent = '0'; totalQuestionsElement.textContent = quizData.length;
    progressBar.style.width = '0%'; 
    
    quizContent.style.display = 'block'; 
    timerContainer.style.display = 'block';
    questionElement.style.display = 'block';
    document.querySelector('.options').style.display = 'grid';
    feedbackElement.style.display = 'block'; 
    prevButton.style.display = 'block';
    nextButton.style.display = 'block';
    nextButton.disabled = false; 
    document.querySelector('.progress').style.display = 'block';
    document.querySelector('.score-container').style.display = 'block';
    
    endQuizElement.style.display = 'none';
    displayQuestion();
}

function displayQuestion() {
    const idx = questionOrder[currentQuestionIndex]; 
    const q = quizData[idx];
    
    questionElement.innerHTML = q.question;
    
    if (q.image_path) {
        const base = "./data/image/"; 
        imageElement.src = q.image_path.startsWith('http') ? q.image_path : base + q.image_path;
        imageElement.style.display = 'block';
    } else {
        imageElement.style.display = 'none';
    }
    
    optionElements.forEach(opt => {
        const k = opt.getAttribute('data-option'); 
        opt.innerHTML = q.options[k]; 
        opt.className = 'option tamil-font';
        
        // ஏற்கனவே பதிலளித்திருந்தால், சரி/தவறு வண்ணங்களைக் காட்டு
        if (userAnswers[idx]) {
            if (k === q.key) opt.classList.add('correct'); 
            else if (userAnswers[idx].selected === k) opt.classList.add('incorrect');
        }
    });
    
    prevButton.disabled = currentQuestionIndex === 0;
    nextButton.textContent = currentQuestionIndex === quizData.length - 1 ? 'முடிவு' : 'அடுத்த கேள்வி';
    progressBar.style.width = `${((currentQuestionIndex + 1) / quizData.length) * 100}%`;
    
    // ✅ புதிய பிழை திருத்தம் இங்கே
    if (userAnswers[idx]) {
        // 1. ஏற்கனவே பதிலளித்த கேள்வி (பின்னால் வரும்போது)
        if (timerInterval) { clearInterval(timerInterval); timerInterval = null; }
        timerContainer.style.display = 'none'; // டைமரை மறைக்கவும்
        nextButton.disabled = false;           // Next பொத்தானை உடனே இயக்கவும்
        
        // முந்தைய பதிலின் ஃபீட்பேக்கைக் காட்டு
        feedbackElement.className = userAnswers[idx].correct ? 'feedback correct' : 'feedback incorrect';
        feedbackElement.innerHTML = userAnswers[idx].correct ? 
            'சரியான பதில்! <i class="fas fa-check-circle"></i>' : 
            `தவறான பதில்! சரியான பதில்: ${q.key}`;
            
    } else {
        // 2. புதிய கேள்வி (முதல் முறையாக வரும்போது)
        timerContainer.style.display = 'block';
        feedbackElement.className = 'feedback'; 
        feedbackElement.textContent = '';
        nextButton.disabled = true; // பதிலளிக்கும் வரை Next முடக்கப்பட்டிருக்கும்
        
        if (timerInterval) clearInterval(timerInterval);
        timeLeft = 30; 
        timerElement.textContent = timeLeft; 
        timerContainer.classList.remove('timer-warning'); 
        questionStartTime = Date.now();
        
        timerInterval = setInterval(() => {
            timeLeft--; 
            timerElement.textContent = timeLeft;
            if (timeLeft <= 10) timerContainer.classList.add('timer-warning');
            if (timeLeft <= 0) { 
                clearInterval(timerInterval); 
                handleTimeUp(); 
            }
        }, 1000);
    }
    
    saveActiveQuizState(); // தேர்வு நிலையைச் சேமி
}

function handleTimeUp() {
    const idx = questionOrder[currentQuestionIndex]; const q = quizData[idx];
    if (!userAnswers[idx]) {
        userAnswers[idx] = { selected: null, correct: false, timeMs: Date.now() - questionStartTime };
        optionElements.forEach(o => { if (o.getAttribute('data-option') === q.key) o.classList.add('correct'); });
        feedbackElement.className = 'feedback incorrect'; feedbackElement.innerHTML = 'நேரம் முடிந்தது! சரியான பதில்: ' + q.key;
        recordWrongAnswer(q);
        
        if (!subjectWiseScores[q.subject]) subjectWiseScores[q.subject] = { score: 0, total: 0 };
        subjectWiseScores[q.subject].total++;
    }
    
    if (currentQuestionIndex === quizData.length - 1) {
        setTimeout(endQuiz, 1500); 
    } else {
        nextButton.disabled = false;
    }
    saveActiveQuizState();
}

optionElements.forEach(opt => {
    opt.addEventListener('click', () => {
        const idx = questionOrder[currentQuestionIndex]; if (userAnswers[idx]) return;
        const q = quizData[idx]; if (timerInterval) { clearInterval(timerInterval); timerInterval = null; }
        const sel = opt.getAttribute('data-option'); const isCorrect = sel === q.key;
        userAnswers[idx] = { selected: sel, correct: isCorrect, timeMs: Date.now() - questionStartTime };
        
        if (isCorrect) {
            score++; scoreElement.textContent = score; opt.classList.add('correct');
            feedbackElement.className = 'feedback correct'; feedbackElement.innerHTML = 'சரியான பதில்! <i class="fas fa-check-circle"></i>';
            recordCorrectAnswer(q);
        } else {
            opt.classList.add('incorrect'); document.querySelector(`.option[data-option="${q.key}"]`).classList.add('correct');
            feedbackElement.className = 'feedback incorrect'; feedbackElement.innerHTML = `தவறான பதில்! சரியான பதில்: ${q.key}`;
            recordWrongAnswer(q);
        }
        
        if (!subjectWiseScores[q.subject]) subjectWiseScores[q.subject] = { score: 0, total: 0 };
        subjectWiseScores[q.subject].total++; if (isCorrect) subjectWiseScores[q.subject].score++;
        nextButton.disabled = false;
        saveActiveQuizState(); // ✅ பதிலளித்த பிறகு State சேமி
    });
});

async function recordWrongAnswer(q) {
    if (!q.id) return; 
    const next = applyAnswer(wrongQuestions[q.id], false); 
    wrongQuestions[q.id] = next;
    const data = { ...next, subject: q.subject, year: q.year, type: q.type, lastAttemptAt: new Date().toISOString() };
    
    saveWQLocal(currentUid, wrongQuestions);
    await syncWQToCloud(currentUid, q.id, data);
}

async function recordCorrectAnswer(q) {
    if (!q.id || !wrongQuestions[q.id]) return; 
    const next = applyAnswer(wrongQuestions[q.id], true); 
    wrongQuestions[q.id] = next;
    const data = { ...next, subject: q.subject, year: q.year, type: q.type, lastAttemptAt: new Date().toISOString() };
    
    saveWQLocal(currentUid, wrongQuestions);
    await syncWQToCloud(currentUid, q.id, data);
}

function nextQuestion() { 
    nextButton.disabled = true;
    if (currentQuestionIndex < quizData.length - 1) { 
        currentQuestionIndex++; 
        displayQuestion(); 
    } else {
        endQuiz(); 
    }
}
function prevQuestion() { 
    if (currentQuestionIndex > 0) { 
        prevButton.disabled = true;
        currentQuestionIndex--; 
        displayQuestion(); 
    } 
}

nextButton.addEventListener('click', nextQuestion); 
prevButton.addEventListener('click', prevQuestion);

// 🔥 End Quiz (Local Save + Cloud Sync)
async function endQuiz() {
    if (timerInterval) { clearInterval(timerInterval); timerInterval = null; }
    clearActiveQuizState(); // ✅ தேர்வு முடிந்ததும் State நீக்கு
    
    const endTime = Date.now();
    const totalTimeMs = endTime - sessionStartTime;
    
    timerContainer.style.display = 'none';
    questionElement.style.display = 'none';
    imageElement.style.display = 'none';
    document.querySelector('.options').style.display = 'none';
    feedbackElement.style.display = 'none';
    prevButton.style.display = 'none';
    nextButton.style.display = 'none';
    document.querySelector('.progress').style.display = 'none';
    document.querySelector('.score-container').style.display = 'none';
    
    endQuizElement.style.display = 'block';
    finalScoreElement.textContent = score;
    finalTotalElement.textContent = quizData.length;
    
    const answered = Object.values(userAnswers);
    const avgTime = answered.length ? (answered.reduce((s, a) => s + (a.timeMs || 0), 0) / answered.length / 1000).toFixed(1) : 0;
    const attentionSpan = document.getElementById('avg-time');
    if (attentionSpan) attentionSpan.textContent = avgTime;
    
    const wrongList = Object.entries(wrongQuestions).filter(([_, q]) => !isMastered(q) && q.weightage >= 2);
    const attentionCount = document.getElementById('attention-count');
    if (attentionCount) attentionCount.textContent = wrongList.length;

    let summaryHTML = '<h3>பாடம் வாரியான மதிப்பெண்:</h3>';
    for (const subject in subjectWiseScores) {
        const perc = (subjectWiseScores[subject].score / subjectWiseScores[subject].total * 100).toFixed(1);
        summaryHTML += `<div class="subject-score"><span>${subject}:</span><span>${perc}% (${subjectWiseScores[subject].score}/${subjectWiseScores[subject].total})</span></div>`;
    }
    scoreSummaryElement.innerHTML = summaryHTML;

    const sessionData = {
        sessionId: 'sess_' + Date.now(),
        device: { autoName: currentDeviceInfo.autoName || "Unknown", customName: currentDeviceInfo.customName || "Unknown" },
        date: new Date().toISOString().slice(0,10), 
        startTime: new Date(sessionStartTime).toISOString(),
        endTime: new Date(endTime).toISOString(),
        totalTimeMs: totalTimeMs, isRepairMode: isRepairMode,
        totalQuestions: quizData.length, correctCount: score,
        answers: quizData.map((q, i) => ({
            questionId: q.id, selected: userAnswers[i]?.selected || null,
            correct: userAnswers[i]?.correct || false, timeMs: userAnswers[i]?.timeMs || 0
        })),
        subjects: Object.entries(subjectWiseScores).map(([name, v]) => ({ name, score: v.score, total: v.total }))
    };

    saveSessionLocal(currentUid, sessionData);
    renderDashboard(loadSessionsLocal(currentUid));
    await syncSessionToCloud(currentUid, sessionData.sessionId, sessionData);

    if (isLocked(wrongQuestions)) {
        setTimeout(() => { showLockScreen(); }, 3000);
    }
}

// Buttons
restartButton.addEventListener('click', () => { showFilterScreen(); });
exitButton.addEventListener('click', () => { showFilterScreen(); });

// 🔥 பழுதுபார்ப்பு தேர்வு பொத்தான்
document.getElementById('repair-btn').addEventListener('click', () => {
    const lockedIds = new Set(Object.entries(wrongQuestions)
        .filter(([_, q]) => q.weightage >= 3 && !isMastered(q)).map(([id]) => id));
    
    if (allQuestions.length === 0) { alert('கேள்விகள் இன்னும் லோட் ஆகவில்லை.'); return; }

    const repairQs = allQuestions.filter(q => lockedIds.has(q.id));
    if (!repairQs.length) { alert('பழுதுபார்ப்பு கேள்விகள் இல்லை'); return; }
    
    isRepairMode = true;
    quizData = shuffleOptions(repairQs); // ✅ இங்கேயும் விருப்பங்களைக் கலைக்கவும்
    
    hideAllScreens();
    userInfoSection.classList.remove('hidden');
    startQuiz(); 
});

applyFiltersBtn.addEventListener('click', applyFilters);
resetFiltersBtn.addEventListener('click', resetFilters);