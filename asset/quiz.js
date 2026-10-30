// Quiz v18.3 - Updated for new Excel Columns
// Developed by GOWTHAMAN B @ 02/09/2025

let allQuestions = [];
let filteredQuestions = [];
let currentQuestionIndex = 0;
let score = 0;
let timer;
let timeLeft = 30;
let userName = "";
let userScores = JSON.parse(localStorage.getItem('drbScores')) || {};
let selectedYear = "all";
let selectedSubject = "all";
let selectedType = "all";

// DOM Elements
const loginSection = document.getElementById('login-section');
const userNameInput = document.getElementById('user-name');
const loginBtn = document.getElementById('login-btn');
const userInfoSection = document.getElementById('user-info-section');
const welcomeMessage = document.getElementById('welcome-message');
const logoutBtn = document.getElementById('logout-btn');
const description = document.getElementById('description');
const instructions = document.getElementById('instructions');
const filterContainer = document.getElementById('filter-container');
const quizContent = document.getElementById('quiz-content');
const questionElement = document.getElementById('question');
const questionImage = document.getElementById('questionImage');
const options = document.querySelectorAll('.option');
const feedback = document.getElementById('feedback');
const prevBtn = document.getElementById('prev-btn');
const nextBtn = document.getElementById('next-btn');
const progressBar = document.getElementById('progress-bar');
const scoreElement = document.getElementById('score');
const totalQuestionsElement = document.getElementById('total-questions');
const timerElement = document.getElementById('timer');
const endQuizDiv = document.getElementById('end-quiz');
const finalScoreElement = document.getElementById('final-score');
const finalTotalElement = document.getElementById('final-total');
const scoreSummaryDiv = document.getElementById('score-summary');
const restartBtn = document.getElementById('restart-btn');
const exitBtn = document.getElementById('exit-btn');
const averageScoreElement = document.getElementById('average-score');
const scoreDetailsElement = document.getElementById('score-details');
const mainScoreBoard = document.getElementById('main-score-board');

// --- LOGIN LOGIC ---
loginBtn.addEventListener('click', () => {
    const name = userNameInput.value.trim();
    if (name) {
        userName = name;
        localStorage.setItem('drbUserName', userName);
        initUser();
    }
});

function initUser() {
    loginSection.classList.add('hidden');
    welcomeMessage.innerText = `வணக்கம், ${userName}!`;
    userInfoSection.classList.remove('hidden');
    description.classList.remove('hidden');
    instructions.classList.remove('hidden');
    filterContainer.classList.remove('hidden');
    loadFile(); // Auto load the Excel file
    updateAverageScore();
}

function updateAverageScore() {
    if (userScores[userName] && userScores[userName].length > 0) {
        let total = userScores[userName].reduce((acc, val) => acc + val, 0);
        let avg = (total / userScores[userName].length).toFixed(1);
        averageScoreElement.innerText = avg;
    } else {
        averageScoreElement.innerText = "0";
    }
}

logoutBtn.addEventListener('click', () => {
    localStorage.removeItem('drbUserName');
    location.reload();
});

// Check if user is already logged in
const savedUser = localStorage.getItem('drbUserName');
if (savedUser) {
    userName = savedUser;
    userNameInput.value = savedUser;
    initUser();
}

// --- EXCEL LOAD & PARSE LOGIC ---
function loadFile() {
    // Automatically try to load the xlsx file (Assuming it is in the same directory)
    fetch('./drb_2025.xlsx')
        .then(response => {
            if (!response.ok) throw new Error("File not found");
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON with empty cells as ""
            const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
            processData(json);
        })
        .catch(error => {
            console.error("Auto-load failed. Please ensure drb_2025.xlsx is in the same folder.", error);
            // Fallback: Allow manual upload if auto-load fails
            alert("Excel file not found! Please place 'drb_2025.xlsx' in the same folder as index.html");
        });
}

function processData(data) {
    // Filter out rows without a question
    allQuestions = data.filter(q => q.question && q.key);
    
    // Populate Filters
    populateFilters();
    
    // Set initial total
    totalQuestionsElement.innerText = allQuestions.length;
}

// --- FILTER LOGIC ---
function populateFilters() {
    const years = new Set();
    const subjects = new Set();
    const types = new Set();

    allQuestions.forEach(q => {
        // Extract Year from 'date' column (e.g., 2025-01-01 00:00:00)
        if (q.date) {
            const yearMatch = String(q.date).match(/\d{4}/);
            if (yearMatch) years.add(yearMatch[0]);
        }
        
        // Get Subject and Type
        if (q.subject) subjects.add(q.subject);
        if (q.type) types.add(q.type);
    });

    // Populate Year Buttons
    const yearContainer = document.getElementById('year-buttons');
    Array.from(years).sort().forEach(year => {
        const btn = document.createElement('button');
        btn.className = 'filter-btn';
        btn.dataset.value = year;
        btn.innerText = year;
        btn.onclick = () => toggleFilter(btn, 'year');
        yearContainer.appendChild(btn);
    });

    // Populate Subject Buttons
    const subjectContainer = document.getElementById('subject-buttons');
    Array.from(subjects).sort().forEach(sub => {
        const btn = document.createElement('button');
        btn.className = 'filter-btn';
        btn.dataset.value = sub;
        btn.innerText = sub;
        btn.onclick = () => toggleFilter(btn, 'subject');
        subjectContainer.appendChild(btn);
    });

    // Populate Type Buttons
    const typeContainer = document.getElementById('type-buttons');
    Array.from(types).sort().forEach(type => {
        const btn = document.createElement('button');
        btn.className = 'filter-btn';
        btn.dataset.value = type;
        btn.innerText = type;
        btn.onclick = () => toggleFilter(btn, 'type');
        typeContainer.appendChild(btn);
    });
}

function toggleFilter(btn, filterType) {
    const parent = btn.parentElement;
    const buttons = parent.querySelectorAll('.filter-btn');
    const value = btn.dataset.value;

    // Remove active from all
    buttons.forEach(b => b.classList.remove('active'));

    // If clicking "Select All", reset
    if (value === 'all') {
        btn.classList.add('active');
        if (filterType === 'year') selectedYear = "all";
        if (filterType === 'subject') selectedSubject = "all";
        if (filterType === 'type') selectedType = "all";
    } else {
        // Toggle specific
        btn.classList.add('active');
        if (filterType === 'year') selectedYear = value;
        if (filterType === 'subject') selectedSubject = value;
        if (filterType === 'type') selectedType = value;
    }
}

document.getElementById('apply-filters-btn').addEventListener('click', () => {
    applyFilters();
    startQuiz();
});

document.getElementById('reset-filters-btn').addEventListener('click', () => {
    selectedYear = "all";
    selectedSubject = "all";
    selectedType = "all";
    
    // Reset UI
    document.querySelectorAll('.filter-section').forEach(section => {
        const buttons = section.querySelectorAll('.filter-btn');
        buttons.forEach(b => b.classList.remove('active'));
        buttons[0].classList.add('active'); // 'All' button
    });
});

function applyFilters() {
    filteredQuestions = allQuestions.filter(q => {
        let matchYear = true;
        let matchSubject = true;
        let matchType = true;

        if (selectedYear !== "all" && q.date) {
            matchYear = String(q.date).includes(selectedYear);
        }
        if (selectedSubject !== "all") {
            matchSubject = q.subject === selectedSubject;
        }
        if (selectedType !== "all") {
            matchType = q.type === selectedType;
        }

        return matchYear && matchSubject && matchType;
    });

    // Shuffle if checked
    if (document.getElementById('randomize-checkbox').checked) {
        filteredQuestions = filteredQuestions.sort(() => Math.random() - 0.5);
    }
}

// --- QUIZ LOGIC ---
function startQuiz() {
    if (filteredQuestions.length === 0) {
        alert("தேர்ந்தெடுக்கப்பட்ட வடிகட்டிகளுக்கு கேள்விகள் இல்லை! (No questions match the filters)");
        return;
    }

    currentQuestionIndex = 0;
    score = 0;
    quizContent.style.display = 'block';
    endQuizDiv.style.display = 'none';
    document.querySelector('.options').style.display = 'flex';
    document.querySelector('.navigation').style.display = 'flex';
    document.querySelector('.progress').style.display = 'flex';
    scoreElement.innerText = score;
    totalQuestionsElement.innerText = filteredQuestions.length;
    
    loadQuestion();
}

function loadQuestion() {
    resetState();
    const q = filteredQuestions[currentQuestionIndex];
    
    // Set Question
    questionElement.innerText = q.question;
    
    // Set Options (a, b, c, d)
    options[0].innerText = q.a || "";
    options[1].innerText = q.b || "";
    options[2].innerText = q.c || "";
    options[3].innerText = q.d || "";
    
    // Set Image
    if (q.image_path) {
        questionImage.src = `./images/${q.image_path}`; // Adjust path if needed
        questionImage.style.display = 'block';
    } else {
        questionImage.style.display = 'none';
    }

    // Update Progress
    progressBar.style.width = `${((currentQuestionIndex + 1) / filteredQuestions.length) * 100}%`;
    
    // Enable/Disable Nav Buttons
    prevBtn.disabled = currentQuestionIndex === 0;
    nextBtn.disabled = false;

    startTimer();
}

function resetState() {
    clearInterval(timer);
    timeLeft = 30;
    timerElement.innerText = timeLeft;
    feedback.innerText = "";
    feedback.className = "feedback";
    options.forEach(opt => {
        opt.classList.remove('correct', 'wrong', 'selected');
        opt.style.pointerEvents = 'auto';
    });
}

function startTimer() {
    timer = setInterval(() => {
        timeLeft--;
        timerElement.innerText = timeLeft;
        if (timeLeft <= 0) {
            clearInterval(timer);
            revealAnswer();
            // Auto move next after 2 seconds
            setTimeout(() => {
                if (currentQuestionIndex < filteredQuestions.length - 1) {
                    currentQuestionIndex++;
                    loadQuestion();
                } else {
                    endQuiz();
                }
            }, 2000);
        }
    }, 1000);
}

function checkAnswer(selectedOption) {
    const q = filteredQuestions[currentQuestionIndex];
    const correctKey = q.key.trim().toUpperCase();
    const selectedKey = selectedOption.dataset.option;
    
    clearInterval(timer);
    options.forEach(opt => opt.style.pointerEvents = 'none');

    // Highlight correct
    options.forEach(opt => {
        if (opt.dataset.option === correctKey) {
            opt.classList.add('correct');
        }
    });

    if (selectedKey === correctKey) {
        score++;
        scoreElement.innerText = score;
        feedback.innerText = "சரியான பதில்! (Correct!)";
        feedback.className = "feedback correct-feedback";
    } else {
        feedback.innerText = `தவறு! சரியான பதில்: ${correctKey}`;
        feedback.className = "feedback wrong-feedback";
        options.forEach(opt => {
            if (opt.dataset.option === selectedKey) {
                opt.classList.add('wrong');
            }
        });
    }
}

options.forEach(opt => {
    opt.addEventListener('click', () => checkAnswer(opt));
});

function revealAnswer() {
    const q = filteredQuestions[currentQuestionIndex];
    const correctKey = q.key.trim().toUpperCase();
    options.forEach(opt => {
        if (opt.dataset.option === correctKey) {
            opt.classList.add('correct');
        }
    });
    feedback.innerText = "நேரம் முடிந்தது!";
    feedback.className = "feedback wrong-feedback";
}

prevBtn.addEventListener('click', () => {
    if (currentQuestionIndex > 0) {
        currentQuestionIndex--;
        loadQuestion();
    }
});

nextBtn.addEventListener('click', () => {
    if (currentQuestionIndex < filteredQuestions.length - 1) {
        currentQuestionIndex++;
        loadQuestion();
    } else {
        endQuiz();
    }
});

function endQuiz() {
    clearInterval(timer);
    quizContent.style.display = 'none';
    endQuizDiv.style.display = 'block';
    
    finalScoreElement.innerText = score;
    finalTotalElement.innerText = filteredQuestions.length;

    // Subject-wise summary
    generateScoreSummary();

    // Save score
    if (!userScores[userName]) userScores[userName] = [];
    const percentage = (score / filteredQuestions.length) * 100;
    userScores[userName].push(percentage);
    localStorage.setItem('drbScores', JSON.stringify(userScores));
    updateAverageScore();
}

function generateScoreSummary() {
    scoreSummaryDiv.innerHTML = "<h3>பாடம் வாரியான மதிப்பெண்:</h3>";
    const subjectScores = {};
    
    // Initialize
    filteredQuestions.forEach(q => {
        const sub = q.subject || "General";
        if (!subjectScores[sub]) subjectScores[sub] = { correct: 0, total: 0 };
        subjectScores[sub].total++;
    });

    // Calculate (This is an approximation since we don't track per-question answers in this simplified version)
    // To do this perfectly, we would need to store an array of answered questions. 
    // For now, let's just show the total questions per subject.
    for (const [sub, data] of Object.entries(subjectScores)) {
        const p = document.createElement('p');
        p.innerText = `${sub}: ${data.total} கேள்விகள்`;
        scoreSummaryDiv.appendChild(p);
    }
}

restartBtn.addEventListener('click', () => {
    startQuiz();
});

exitBtn.addEventListener('click', () => {
    endQuizDiv.style.display = 'none';
    filterContainer.classList.remove('hidden');
    quizContent.style.display = 'none';
});