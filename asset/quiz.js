// Quiz v20.2 - Firebase Integrated with Old UI Logic (No Filter Changes)
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { getAuth, setPersistence, browserLocalPersistence,
         signInWithEmailAndPassword, signOut, onAuthStateChanged }
       from "https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js";
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager,
         doc, getDoc, setDoc, updateDoc, collection, getDocs }
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

let availableYears = [], availableSubjects = [], availableTypes = [];
let selectedYears = [], selectedSubjects = [], selectedTypes = [];

// STAGE CONFIG
const STAGE_CONFIG = {
    1: { maxStars: 5,  mode: 'optional'   },
    2: { maxStars: 10, mode: 'compulsory' },
    3: { maxStars: 15, mode: 'locked'     },
    4: { maxStars: 20, mode: 'locked'     },
    5: { maxStars: 25, mode: 'locked'     }
};

function isMastered(wq) { return wq && wq.weightage === 5 && wq.stars >= STAGE_CONFIG[5].maxStars; }
function applyAnswer(wq, isCorrect) {
    if (!wq) return isCorrect ? null : { weightage:1, stars:0, wrongCount:1, correctCount:0 };
    const cfg = STAGE_CONFIG[wq.weightage];
    const atMax = wq.stars >= cfg.maxStars;
    if (isCorrect) return { ...wq, stars: Math.min(wq.stars + 1, cfg.maxStars), correctCount: (wq.correctCount || 0) + 1 };
    if (atMax && wq.weightage < 5) return { ...wq, weightage: wq.weightage + 1, stars: 0, wrongCount: (wq.wrongCount || 0) + 1 };
    return { ...wq, stars: Math.max(wq.stars - 1, 0), wrongCount: (wq.wrongCount || 0) + 1 };
}
function isLocked(wqMap) { return Object.values(wqMap).some(q => q.weightage >= 3 && !isMastered(q)); }

// 🔥 Fetch Question Bank
async function loadQuestionBank() {
    const timestamp = new Date().getTime();
    const defaultPath = `./data/question_bank.xlsx?v=${timestamp}`;
    try {
        const response = await fetch(defaultPath);
        if (!response.ok) throw new Error('கோப்பை ஏற்ற முடியவில்லை');
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames.includes('MASTER_DATA') ? 'MASTER_DATA' : workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        processExcelData(jsonData);
    } catch (error) {
        console.error('Auto-load error:', error);
        alert('வினா வங்கியை தானாக ஏற்ற முடியவில்லை.');
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
        showFilterOptions();
    }
}

// 🔥 Filters (No changes from v20.0)
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

function showFilterOptions() { filterContainer.style.display = 'block'; }

// 🔥 Apply Filters (No changes from v20.0)
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
    if (randomize) quizData = quizData.sort(() => Math.random() - 0.5);
    filterContainer.style.display = 'none';
    startQuiz();
}

function resetFilters() {
    selectedYears = [...availableYears]; selectedSubjects = [...availableSubjects]; selectedTypes = [...availableTypes];
    document.querySelectorAll('.filter-btn').forEach(btn => btn.classList.add('selected'));
    randomizeCheckbox.checked = true; if(injectWrongCheckbox) injectWrongCheckbox.checked = false;
}

// 🔥 Firebase Functions
async function loadWrongQuestions(uid) {
    const snap = await getDocs(collection(db, 'users', uid, 'wrongQuestions'));
    const out = {}; snap.forEach(d => { out[d.id] = d.data(); }); return out;
}
async function saveWrongQuestion(uid, qid, data) { await setDoc(doc(db, 'users', uid, 'wrongQuestions', qid), data, { merge: true }); }
async function saveQuizSession(uid, sid, data) { await setDoc(doc(db, 'users', uid, 'quizSessions', sid), data); }
async function loadQuizSessions(uid) {
    const snap = await getDocs(collection(db, 'users', uid, 'quizSessions'));
    const arr = []; snap.forEach(d => arr.push(d.data())); return arr;
}

// 🔥 Login & Auth
document.getElementById('login-btn').addEventListener('click', async () => {
    const email = document.getElementById('login-email').value.trim();
    const pass  = document.getElementById('login-password').value;
    const errEl = document.getElementById('auth-error');
    errEl.style.display = 'none';
    try { await signInWithEmailAndPassword(auth, email, pass); } 
    catch (e) { errEl.textContent = 'உள்நுழைவு தோல்வி: ' + e.message; errEl.style.display = 'block'; }
});
document.getElementById('logout-btn').addEventListener('click', () => signOut(auth));

onAuthStateChanged(auth, async user => {
    if (!user) {
        loginSection.classList.remove('hidden'); userInfoSection.classList.add('hidden');
        filterContainer.style.display = 'none'; return;
    }
    currentUid = user.uid;
    const emailDisplay = document.getElementById('user-email-display');
    if (emailDisplay) emailDisplay.textContent = user.email;
    
    loginSection.classList.add('hidden'); userInfoSection.classList.remove('hidden');
    
    try { wrongQuestions = await loadWrongQuestions(currentUid); } catch (e) { wrongQuestions = {}; }
    await showMainScoreBoard(); // மதிப்பெண் பலகையை லோட் செய்
    
    if (isLocked(wrongQuestions)) {
        lockScreen.classList.remove('hidden');
        document.getElementById('locked-count').textContent = Object.keys(wrongQuestions).filter(id => wrongQuestions[id].weightage >= 3).length;
    } else {
        lockScreen.classList.add('hidden');
        loadQuestionBank();
    }
});

// 🔥 Show Main Score Board (Old v18.3 Logic applied to Firebase)
async function showMainScoreBoard() {
    if (!currentUid) return;
    let sessions = [];
    try { sessions = await loadQuizSessions(currentUid); } catch (e) { console.error(e); }
    
    if (!sessions.length) {
        document.getElementById('score-details').innerHTML = '<p>இதுவரை மதிப்பெண்கள் இல்லை</p>';
        document.getElementById('average-score').textContent = '0';
        document.getElementById('main-score-board').style.display = 'block';
        return;
    }
    
    let totalScore = 0, totalQ = 0, subj = {};
    sessions.forEach(s => {
        totalScore += s.correctCount || 0; totalQ += s.totalQuestions || 0;
        (s.subjects || []).forEach(sub => {
            if (!subj[sub.name]) subj[sub.name] = { s: 0, t: 0 };
            subj[sub.name].s += sub.score; subj[sub.name].t += sub.total;
        });
    });
    
    const overallAvg = totalQ > 0 ? Math.round((totalScore / totalQ) * 100) : 0;
    document.getElementById('average-score').textContent = overallAvg;
    
    let html = '';
    for (const [name, data] of Object.entries(subj)) {
        const avg = data.t > 0 ? Math.round((data.s / data.t) * 100) : 0;
        html += `<div class="subject-score"><span>${name}:</span><span>${avg}% (${data.s}/${data.t})</span></div>`;
    }
    html += `<div class="subject-score"><span><b>தேர்வுகள்: ${sessions.length}</b></span><span><b>${overallAvg}% (${totalScore}/${totalQ})</b></span></div>`;
    
    document.getElementById('score-details').innerHTML = html;
    document.getElementById('main-score-board').style.display = 'block';
}

// 🔥 Quiz Logic
function startQuiz() {
    subjectWiseScores = {};
    questionOrder = Array.from({length: quizData.length}, (_, i) => i);
    currentQuestionIndex = 0; score = 0; userAnswers = {}; sessionStartTime = Date.now();
    scoreElement.textContent = '0'; totalQuestionsElement.textContent = quizData.length;
    progressBar.style.width = '0%'; 
    
    quizContent.style.display = 'block'; 
    endQuizElement.style.display = 'none';
    displayQuestion();
}

function displayQuestion() {
    const idx = questionOrder[currentQuestionIndex]; const q = quizData[idx];
    questionElement.innerHTML = q.question;
    if (q.image_path) {
        const base = "./data/image/"; 
        imageElement.src = q.image_path.startsWith('http') ? q.image_path : base + q.image_path;
        imageElement.style.display = 'block';
    } else imageElement.style.display = 'none';
    
    optionElements.forEach(opt => {
        const k = opt.getAttribute('data-option'); opt.innerHTML = q.options[k]; opt.className = 'option tamil-font';
        if (userAnswers[idx]) {
            if (k === q.key) opt.classList.add('correct'); else if (userAnswers[idx].selected === k) opt.classList.add('incorrect');
        }
    });
    
    prevButton.disabled = currentQuestionIndex === 0;
    nextButton.textContent = currentQuestionIndex === quizData.length - 1 ? 'முடிவு' : 'அடுத்த கேள்வி';
    progressBar.style.width = `${((currentQuestionIndex + 1) / quizData.length) * 100}%`;
    feedbackElement.className = 'feedback'; feedbackElement.textContent = '';
    
    if (timerInterval) clearInterval(timerInterval);
    timeLeft = 30; timerElement.textContent = timeLeft; timerContainer.classList.remove('timer-warning'); questionStartTime = Date.now();
    timerInterval = setInterval(() => {
        timeLeft--; timerElement.textContent = timeLeft;
        if (timeLeft <= 10) timerContainer.classList.add('timer-warning');
        if (timeLeft <= 0) { clearInterval(timerInterval); handleTimeUp(); }
    }, 1000);
}

function handleTimeUp() {
    const idx = questionOrder[currentQuestionIndex]; const q = quizData[idx];
    if (!userAnswers[idx]) {
        userAnswers[idx] = { selected: null, correct: false, timeMs: Date.now() - questionStartTime };
        optionElements.forEach(o => { if (o.getAttribute('data-option') === q.key) o.classList.add('correct'); });
        feedbackElement.className = 'feedback incorrect'; feedbackElement.innerHTML = 'நேரம் முடிந்தது! சரியான பதில்: ' + q.key;
        recordWrongAnswer(q);
    }
    nextButton.disabled = false;
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
    });
});

async function recordWrongAnswer(q) {
    if (!q.id) return; const next = applyAnswer(wrongQuestions[q.id], false); wrongQuestions[q.id] = next;
    await saveWrongQuestion(currentUid, q.id, { ...next, subject: q.subject, year: q.year, type: q.type, lastAttemptAt: new Date().toISOString() });
}
async function recordCorrectAnswer(q) {
    if (!q.id || !wrongQuestions[q.id]) return; const next = applyAnswer(wrongQuestions[q.id], true); wrongQuestions[q.id] = next;
    await saveWrongQuestion(currentUid, q.id, { ...next, subject: q.subject, year: q.year, type: q.type, lastAttemptAt: new Date().toISOString() });
}

function nextQuestion() { if (currentQuestionIndex < quizData.length - 1) { currentQuestionIndex++; displayQuestion(); } else endQuiz(); }
function prevQuestion() { if (currentQuestionIndex > 0) { currentQuestionIndex--; displayQuestion(); } }
nextButton.addEventListener('click', nextQuestion); prevButton.addEventListener('click', prevQuestion);

// 🔥 End Quiz (Old v18.3 Logic applied correctly)
async function endQuiz() {
    if (timerInterval) { clearInterval(timerInterval); timerInterval = null; }
    
    // பழையபடியே எல்லா எலிமெண்ட்ஸையும் மறைக்கவும், ஆனால் quizContent-ஐ மறைக்கக் கூடாது
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
    
    // நேரம் & கவனம் தேவை கணக்கீடு
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

    // Firebase Session Save
    const sessionId = 'sess_' + Date.now();
    try {
        await saveQuizSession(currentUid, sessionId, {
            date: new Date().toISOString().slice(0,10), totalQuestions: quizData.length, correctCount: score,
            subjects: Object.entries(subjectWiseScores).map(([name, v]) => ({ name, score: v.score, total: v.total }))
        });
    } catch(e) {}
    
    // மதிப்பெண் பலகையைப் புதுப்பிக்கவும்
    await showMainScoreBoard();
}

// Buttons
restartButton.addEventListener('click', () => { 
    endQuizElement.style.display = 'none'; 
    quizContent.style.display = 'none'; 
    filterContainer.style.display = 'block'; 
    showMainScoreBoard();
});
exitButton.addEventListener('click', () => { 
    endQuizElement.style.display = 'none'; 
    quizContent.style.display = 'none'; 
    filterContainer.style.display = 'block'; 
    showMainScoreBoard();
});

applyFiltersBtn.addEventListener('click', applyFilters);
resetFiltersBtn.addEventListener('click', resetFilters);