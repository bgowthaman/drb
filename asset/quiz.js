// Quiz v18.3 - Restored quiz_bk2.js logic with new Excel Columns (MASTER_DATA)
// Developed by GOWTHAMAN B

// DOM elements
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
const descriptionElement = document.getElementById('description');
const instructionsElement = document.getElementById('instructions');
const timerContainer = document.getElementById('timer-container');
const timerElement = document.getElementById('timer');
const quizContent = document.getElementById('quiz-content');
const scoreSummaryElement = document.getElementById('score-summary');
const imageElement = document.getElementById('questionImage');

// Filter elements
const filterContainer = document.getElementById('filter-container');
const yearButtonsContainer = document.getElementById('year-buttons');
const subjectButtonsContainer = document.getElementById('subject-buttons');
const typeButtonsContainer = document.getElementById('type-buttons');
const randomizeCheckbox = document.getElementById('randomize-checkbox');
const applyFiltersBtn = document.getElementById('apply-filters-btn');
const resetFiltersBtn = document.getElementById('reset-filters-btn');

// User elements
const welcomeMessage = document.getElementById('welcome-message');
const userNameInput = document.getElementById('user-name');
const loginBtn = document.getElementById('login-btn');
const logoutBtn = document.getElementById('logout-btn');
const loginSection = document.getElementById('login-section');
const userInfoSection = document.getElementById('user-info-section');

// Score elements
const mainScoreBoard = document.getElementById('main-score-board');
const scoreDetails = document.getElementById('score-details');
const averageScoreElement = document.getElementById('average-score');

// Quiz state
let quizData = [];
let currentQuestionIndex = 0;
let score = 0;
let selectedOption = null;
let questionOrder = [];
let userAnswers = {}; 
let timerInterval = null;
let timeLeft = 30;		
let allQuestions = []; 
let availableYears = [];
let availableSubjects = [];
let availableTypes = []; 
let selectedYears = [];
let selectedSubjects = [];
let selectedTypes = []; 
let currentFilters = {};
let currentUserName = '';
let subjectWiseScores = {}; 
let subjectWiseTotals = {}; 

// Auto-load question bank from default path with cache busting
function autoLoadQuestionBank() {
    // Add timestamp to prevent browser caching of the old file
    const timestamp = new Date().getTime();
    const defaultPath = `./data/question_bank.xlsx?v=${timestamp}`;
    
    fetch(defaultPath)
        .then(response => {
            if (!response.ok) {
                throw new Error('கோப்பை ஏற்ற முடியவில்லை');
            }
            return response.arrayBuffer();
        })
        .then(arrayBuffer => {
            const data = new Uint8Array(arrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Explicitly target 'MASTER_DATA' sheet
            const sheetName = workbook.SheetNames.includes('MASTER_DATA') ? 'MASTER_DATA' : workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // Convert to JSON objects using column headers
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
            
            // Process the data
            processExcelData(jsonData);
        })
        .catch(error => {
            console.error('Auto-load error:', error);
            alert('வினா வங்கியை தானாக ஏற்ற முடியவில்லை. கோப்பு பாதை: ' + defaultPath);
        });
}

// Initialize the application
function initApp() {
    const savedUserName = localStorage.getItem('quizUserName');
    
    if (savedUserName) {
        currentUserName = savedUserName;
        welcomeMessage.textContent = `வணக்கம் ${savedUserName}!`;
        loginSection.classList.add('hidden');
        userInfoSection.classList.remove('hidden');
        showMainScoreBoard(savedUserName);
        autoLoadQuestionBank();
    } else {
        loginSection.classList.remove('hidden');
        userInfoSection.classList.add('hidden');
    }
}

// Login user
function loginUser() {
    const userName = userNameInput.value.trim();
    if (userName) {
        currentUserName = userName;
        welcomeMessage.textContent = `வணக்கம் ${userName}!`;
        loginSection.classList.add('hidden');
        userInfoSection.classList.remove('hidden');
        localStorage.setItem('quizUserName', userName);
        showMainScoreBoard(userName);
        autoLoadQuestionBank();
    } else {
        alert('தயவு செய்து உங்கள் பெயரை உள்ளிடவும்');
    }
}

// Show main score board
function showMainScoreBoard(userName) {
    const userScores = JSON.parse(localStorage.getItem('userScores') || '{}');
    const userData = userScores[userName] || {};
    
    let scoresHTML = '';
    let totalScore = 0;
    let totalPossible = 0;
    let subjectAverages = {};
    let quizCounts = {}; 
    
    if (Object.keys(userData).length === 0) {
        scoresHTML = '<p>இதுவரை மதிப்பெண்கள் இல்லை</p>';
        averageScoreElement.textContent = '0';
    } else {
        subjectAverages = {};
        quizCounts = {};
        
        for (const [quizId, quizData] of Object.entries(userData)) {
            if (quizId === 'quizzes') continue;
            quizCounts[quizId] = true;
            
            for (const [subject, data] of Object.entries(quizData.subjects || {})) {
                if (!subjectAverages[subject]) {
                    subjectAverages[subject] = { totalScore: 0, totalPossible: 0, count: 0 };
                }
                subjectAverages[subject].totalScore += data.score;
                subjectAverages[subject].totalPossible += data.total;
                subjectAverages[subject].count++;
                
                totalScore += data.score;
                totalPossible += data.total;
            }
        }
        
        const overallAverage = totalPossible > 0 ? Math.round((totalScore / totalPossible) * 100) : 0;
        averageScoreElement.textContent = overallAverage;
        
        scoresHTML += '<div class="average-score"></div>';
        
        for (const [subject, data] of Object.entries(subjectAverages)) {
            const average = data.totalPossible > 0 ? Math.round((data.totalScore / data.totalPossible) * 100) : 0;
            scoresHTML += `
                <div class="subject-score">
                    <span>${subject}:</span>
                    <span>${average}% (${data.totalScore}/${data.totalPossible})</span>
                </div>
            `;
        }
        
        const totalQuizzes = Object.keys(quizCounts).length;
        scoresHTML += `<div class="subject-score">
            <span><b>தேர்வுகள்: ${totalQuizzes}</b></span>
            <span><b>${overallAverage}% (${totalScore}/${totalPossible})</b></span>
        </div>`;
    }
    
    scoreDetails.innerHTML = scoresHTML;
    mainScoreBoard.style.display = 'block';
}

// Save user scores
function saveUserScores() {
    const userScores = JSON.parse(localStorage.getItem('userScores') || '{}');
    if (!userScores[currentUserName]) userScores[currentUserName] = { quizzes: {} };
    const quizId = `quiz_${Date.now()}`;
    
    userScores[currentUserName][quizId] = {
        date: new Date().toISOString(),
        totalScore: score,
        totalQuestions: quizData.length,
        subjects: {...subjectWiseScores}
    };
    
    if (!userScores[currentUserName].quizzes) userScores[currentUserName].quizzes = {};
    userScores[currentUserName].quizzes[quizId] = true;
    
    localStorage.setItem('userScores', JSON.stringify(userScores));
    localStorage.setItem('quizUserName', currentUserName);
}

// Logout user
function logoutUser() {
    localStorage.removeItem('quizUserName');
    currentUserName = '';
    window.location.reload(); 
}

// Process Excel data from object-based parsing
function processExcelData(data) {
    const processText = (text) => {
        if (!text) return "";
        let processed = String(text).replace(/\n/g, '<br>');
        processed = processed.replace(/<br\s*\/?>/gi, '<br>');
        return processed;
    };

    const extractYear = (dateVal) => {
        if (!dateVal) return "வருடம்";
        
        // Excel-ல் தேதி எண்ணாக (Serial Date) இருந்தால்
        if (typeof dateVal === 'number') {
            const date = new Date((dateVal - 25569) * 86400 * 1000);
            return date.getFullYear().toString();
        }
        
        // String-ஆக இருந்தால்
        const match = String(dateVal).match(/\d{4}/);
        return match ? match[0] : "வருடம்";
    };

    quizData = data.map(row => {
        return {
            year: extractYear(row.date), 
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
            key: String(row.key || "A").trim().toUpperCase(),					
        };
    }).filter(item => item.question !== "கேள்வி இல்லை");
    
    if (quizData.length > 0) {
        allQuestions = quizData;
        extractFilters(allQuestions);
        showFilterOptions();
    } else {
        questionElement.textContent = "கோப்பில் ஏற்றத்தக்க தரவு இல்லை";
    }
}

// Extract available filters
function extractFilters(questions) {
    availableYears = ['all'];
    availableSubjects = ['all'];
    availableTypes = ['all'];
    
    questions.forEach(item => {
        if (item.year && !availableYears.includes(item.year)) availableYears.push(item.year);
        if (item.subject && !availableSubjects.includes(item.subject)) availableSubjects.push(item.subject);
        if (item.type && !availableTypes.includes(item.type)) availableTypes.push(item.type);
    });
    
    yearButtonsContainer.innerHTML = '';
    availableYears.sort().forEach(year => {
        const button = document.createElement('button');
        button.className = 'filter-btn';
        button.textContent = year === 'all' ? 'அனைத்தும்' : year;
        button.setAttribute('data-value', year);
        button.addEventListener('click', () => toggleFilter('year', year, button));
        yearButtonsContainer.appendChild(button);
    });
    
    subjectButtonsContainer.innerHTML = '';
    availableSubjects.sort().forEach(subject => {
        const button = document.createElement('button');
        button.className = 'filter-btn';
        button.textContent = subject === 'all' ? 'அனைத்தும்' : subject;
        button.setAttribute('data-value', subject);
        button.addEventListener('click', () => toggleFilter('subject', subject, button));
        subjectButtonsContainer.appendChild(button);
    });
    
    typeButtonsContainer.innerHTML = '';
    availableTypes.sort().forEach(type => {
        const button = document.createElement('button');
        button.className = 'filter-btn';
        button.textContent = type === 'all' ? 'அனைத்தும்' : type;
        button.setAttribute('data-value', type);
        button.addEventListener('click', () => toggleFilter('type', type, button));
        typeButtonsContainer.appendChild(button);
    });
    
    loadSavedFilters();
}

// Toggle filter selection
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
        if (button.classList.contains('selected')) selectedYears.push(value);
        else selectedYears = selectedYears.filter(y => y !== value);
        
        const allBtn = document.querySelector('#year-buttons .filter-btn[data-value="all"]');
        if (selectedYears.length === availableYears.length - 1) {
            allBtn.classList.add('selected');
            selectedYears.push('all');
        } else {
            allBtn.classList.remove('selected');
            selectedYears = selectedYears.filter(y => y !== 'all');
        }
    } else if (type === 'subject') {
        if (button.classList.contains('selected')) selectedSubjects.push(value);
        else selectedSubjects = selectedSubjects.filter(s => s !== value);
        
        const allBtn = document.querySelector('#subject-buttons .filter-btn[data-value="all"]');
        if (selectedSubjects.length === availableSubjects.length - 1) {
            allBtn.classList.add('selected');
            selectedSubjects.push('all');
        } else {
            allBtn.classList.remove('selected');
            selectedSubjects = selectedSubjects.filter(s => s !== 'all');
        }
    } else {
        if (button.classList.contains('selected')) selectedTypes.push(value);
        else selectedTypes = selectedTypes.filter(t => t !== value);
        
        const allBtn = document.querySelector('#type-buttons .filter-btn[data-value="all"]');
        if (selectedTypes.length === availableTypes.length - 1) {
            allBtn.classList.add('selected');
            selectedTypes.push('all');
        } else {
            allBtn.classList.remove('selected');
            selectedTypes = selectedTypes.filter(t => t !== 'all');
        }
    }
}

// Show filter options
function showFilterOptions() {
    if (descriptionElement) descriptionElement.classList.add('hidden');
    if (instructionsElement) instructionsElement.classList.add('hidden');
    filterContainer.style.display = 'block';
}

// Apply filters and start quiz
function applyFilters() {
    const randomize = randomizeCheckbox.checked;
    saveFilters(selectedYears, selectedSubjects, selectedTypes, randomize);
    
    let filteredQuestions = allQuestions;
    
    if (selectedYears.length > 0 && !selectedYears.includes('all')) {
        filteredQuestions = filteredQuestions.filter(q => selectedYears.includes(q.year));
    }
    if (selectedSubjects.length > 0 && !selectedSubjects.includes('all')) {
        filteredQuestions = filteredQuestions.filter(q => selectedSubjects.includes(q.subject));
    }
    if (selectedTypes.length > 0 && !selectedTypes.includes('all')) {
        filteredQuestions = filteredQuestions.filter(q => selectedTypes.includes(q.type));
    }
    
    if (filteredQuestions.length === 0) {
        alert('தேர்ந்தெடுத்த வடிப்பான்களுக்கு ஏற்ற கேள்விகள் இல்லை.');
        return;
    }
    
    currentFilters = {
        years: [...selectedYears],
        subjects: [...selectedSubjects],
        types: [...selectedTypes],
        randomize: randomize
    };
    
    quizData = filteredQuestions.map(item => {
        const processText = (text) => {
            if (!text) return "";
            return String(text).replace(/\n/g, '<br>').replace(/<br\s*\/?>/gi, '<br>');
        };
        return {
            question: processText(item.question) || "கேள்வி இல்லை",
            image_path: item.image_path || "",
            options: {
                A: processText(item.options?.A) || "விருப்பம் A",
                B: processText(item.options?.B) || "விருப்பம் B",
                C: processText(item.options?.C) || "விருப்பம் C",
                D: processText(item.options?.D) || "விருப்பம் D"
            },
            key: String(item.key || "A").trim().toUpperCase(),
            year: item.year || "வருடம்",
            subject: item.subject || "பொது",
            type: item.type || "வடிவம்"
        };
    }).filter(item => item.question !== "கேள்வி இல்லை");
    
    if (quizData.length > 0) {
        filterContainer.style.display = 'none';
        startQuiz(randomize);
    }
}

// Save/Load/Reset Filters
function saveFilters(years, subjects, types, randomize) {
    if (typeof(Storage) !== "undefined") {
        localStorage.setItem('quizFilters', JSON.stringify({ years, subjects, types, randomize }));
    }
}

function loadSavedFilters() {
    if (typeof(Storage) !== "undefined") {
        const savedFilters = localStorage.getItem('quizFilters');
        if (savedFilters) {
            const filters = JSON.parse(savedFilters);
            if (filters.years) {
                selectedYears = filters.years;
                document.querySelectorAll('#year-buttons .filter-btn').forEach(btn => {
                    if (selectedYears.includes(btn.getAttribute('data-value'))) btn.classList.add('selected');
                });
            }
            if (filters.subjects) {
                selectedSubjects = filters.subjects;
                document.querySelectorAll('#subject-buttons .filter-btn').forEach(btn => {
                    if (selectedSubjects.includes(btn.getAttribute('data-value'))) btn.classList.add('selected');
                });
            }
            if (filters.types) {
                selectedTypes = filters.types;
                document.querySelectorAll('#type-buttons .filter-btn').forEach(btn => {
                    if (selectedTypes.includes(btn.getAttribute('data-value'))) btn.classList.add('selected');
                });
            }
            if (filters.randomize !== undefined) randomizeCheckbox.checked = filters.randomize;
        } else {
            selectedYears = [...availableYears];
            selectedSubjects = [...availableSubjects];
            selectedTypes = [...availableTypes];
            document.querySelectorAll('.filter-btn').forEach(btn => btn.classList.add('selected'));
        }
    }
}

function resetFilters() {
    selectedYears = [...availableYears];
    selectedSubjects = [...availableSubjects];
    selectedTypes = [...availableTypes];
    document.querySelectorAll('.filter-btn').forEach(btn => btn.classList.add('selected'));
    randomizeCheckbox.checked = true;
}

// Start Quiz
function startQuiz(randomize = true) {
    subjectWiseScores = {};
    subjectWiseTotals = {};			
    questionOrder = Array.from({length: quizData.length}, (_, i) => i);
    
    currentQuestionIndex = 0;
    score = 0;
    selectedOption = null;
    userAnswers = {};
    
    scoreElement.textContent = score;
    totalQuestionsElement.textContent = quizData.length;
    progressBar.style.width = '0%';
    
    quizContent.style.display = 'block';
    endQuizElement.style.display = 'none';
    
    displayQuestion();
}

// Display Question Image
function displayQuestionImage() {
    const questionIndex = questionOrder[currentQuestionIndex];
    const question = quizData[questionIndex];
    
    if (question.image_path && question.image_path !== "") {
        const baseImagePath = "./data/image/"; // Change this if your images are elsewhere
        let fullImagePath = question.image_path;
        if(!fullImagePath.startsWith('http') && !fullImagePath.startsWith('./') && !fullImagePath.startsWith('images/')){
            fullImagePath = baseImagePath + question.image_path;
        }
        
        imageElement.src = fullImagePath; 
        imageElement.style.display = 'block';
        imageElement.alt = "கேள்விக்குரிய படம்";
        imageElement.onerror = function() { imageElement.style.display = 'none'; };
    } else {
        imageElement.style.display = 'none';
    }
}

// Display Question
function displayQuestion() {
    const questionIndex = questionOrder[currentQuestionIndex];
    const question = quizData[questionIndex];
    
    questionElement.innerHTML = question.question;
    displayQuestionImage();
    
    optionElements.forEach(option => {
        const optionKey = option.getAttribute('data-option');
        option.innerHTML = question.options[optionKey];
        option.className = 'option tamil-font';
        
        if (userAnswers[questionIndex]) {
            if (optionKey === question.key) {
                option.classList.add('correct');
            } else if (userAnswers[questionIndex].selected === optionKey && userAnswers[questionIndex].selected !== question.key) {
                option.classList.add('incorrect');
            }
            if (userAnswers[questionIndex].selected === optionKey) {
                option.classList.add('selected');
            }                    
        }
    });
                                
    prevButton.disabled = currentQuestionIndex === 0;
    nextButton.disabled = currentQuestionIndex === quizData.length;
    nextButton.textContent = currentQuestionIndex === quizData.length - 1 ? 'முடிவு' : 'அடுத்த கேள்வி';
    progressBar.style.width = `${((currentQuestionIndex + 1) / quizData.length) * 100}%`;
    
    feedbackElement.className = 'feedback';
    feedbackElement.textContent = '';
    startTimer();
}

// Timer
function startTimer() {
    if (timerInterval) clearInterval(timerInterval);
    timeLeft = 30;
    timerElement.textContent = timeLeft;
    timerContainer.classList.remove('timer-warning');
    
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

function handleTimeUp() {
    const questionIndex = questionOrder[currentQuestionIndex];
    const question = quizData[questionIndex];
    
    if (!userAnswers[questionIndex]) {
        userAnswers[questionIndex] = { selected: null, correct: false, score: 0 };
        optionElements.forEach(option => {
            if (option.getAttribute('data-option') === question.key) option.classList.add('correct');
        });
        feedbackElement.className = 'feedback incorrect';
        feedbackElement.innerHTML = 'நேரம் முடிந்துவிட்டது! சரியான பதில்: ' + question.key;
    }
    nextButton.disabled = false;
}

// Score Tracking
function updateSubjectScore(subject, isCorrect) {
    if (!subjectWiseScores[subject]) subjectWiseScores[subject] = { score: 0, total: 0 };
    if (!subjectWiseTotals[subject]) subjectWiseTotals[subject] = 0;
    subjectWiseScores[subject].total++;
    subjectWiseTotals[subject]++;
    if (isCorrect) subjectWiseScores[subject].score++;
}

// Check Answer
function checkAnswer(selectedOptionKey) {
    const questionIndex = questionOrder[currentQuestionIndex];
    const question = quizData[questionIndex];
    
    if (timerInterval) { clearInterval(timerInterval); timerInterval = null; }
    
    const correctAnswer = question.key;
    const isCorrect = selectedOptionKey === correctAnswer;
    
    userAnswers[questionIndex] = { selected: selectedOptionKey, correct: isCorrect, score: isCorrect ? 1 : 0 };
    
    if (isCorrect) score++;
    scoreElement.textContent = score;
    
    feedbackElement.className = isCorrect ? 'feedback correct' : 'feedback incorrect';
    feedbackElement.innerHTML = isCorrect ? 
        'சரியான பதில்! <i class="fas fa-check-circle"></i>' : 
        'தவறான பதில்! சரியான பதில்: ' + correctAnswer;
    
    optionElements.forEach(option => {
        const optionKey = option.getAttribute('data-option');
        if (optionKey === correctAnswer) option.classList.add('correct');
        else if (optionKey === selectedOptionKey && !isCorrect) option.classList.add('incorrect');
        if (optionKey === selectedOptionKey) option.classList.add('selected');
    });
    
    updateSubjectScore(question.subject, isCorrect);
    nextButton.disabled = false;
}

// Navigation
function nextQuestion() {
    if (currentQuestionIndex < quizData.length - 1) {
        currentQuestionIndex++;
        selectedOption = null;
        displayQuestion();
    } else {
        endQuiz();
    }
}

function prevQuestion() {
    if (currentQuestionIndex > 0) {
        currentQuestionIndex--;
        selectedOption = null;
        displayQuestion();
    }
}

// End Quiz
function endQuiz() {
    if (timerInterval) { clearInterval(timerInterval); timerInterval = null; }
    
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
    
    let summaryHTML = '<h3>பாடம் வாரியான மதிப்பெண்:</h3>';
    for (const subject in subjectWiseScores) {
        const subjectScore = subjectWiseScores[subject].score;
        const subjectTotal = subjectWiseScores[subject].total;
        const percentage = subjectTotal > 0 ? (subjectScore / subjectTotal * 100).toFixed(1) : 0;
        summaryHTML += `
            <div class="subject-score">
                <span>${subject}:</span>
                <span>${percentage}% (${subjectScore}/${subjectTotal})</span>
            </div>
        `;
    }
    scoreSummaryElement.innerHTML = summaryHTML;
    saveUserScores();
}

// Restart / Exit
function restartQuiz() {
    timerContainer.style.display = 'block';
    questionElement.style.display = 'block';
    document.querySelector('.options').style.display = 'grid';
    feedbackElement.style.display = 'none';
    prevButton.style.display = 'block';
    nextButton.style.display = 'block';
    document.querySelector('.progress').style.display = 'block';
    document.querySelector('.score-container').style.display = 'block';
    endQuizElement.style.display = 'none';
    
    if (Object.keys(currentFilters).length > 0) applyFilters();
    else startQuiz();
}

function exitQuiz() {
    if (timerInterval) clearInterval(timerInterval);
    quizData = [];
    currentQuestionIndex = 0;
    score = 0;
    selectedOption = null;
    userAnswers = {};
    
    quizContent.style.display = 'none';
    endQuizElement.style.display = 'none';
    autoLoadQuestionBank();	
}

// Option Click Listeners
optionElements.forEach(option => {
    option.addEventListener('click', () => {
        if (userAnswers[questionOrder[currentQuestionIndex]] !== undefined) return;
        
        optionElements.forEach(opt => opt.classList.remove('selected'));
        option.classList.add('selected');
        selectedOption = option.getAttribute('data-option');
        
        const question = quizData[questionOrder[currentQuestionIndex]];
        const isCorrect = selectedOption === question.key;
        
        userAnswers[questionOrder[currentQuestionIndex]] = { selected: selectedOption, correct: isCorrect };
        
        if (isCorrect) {
            score++;
            scoreElement.textContent = score;
            option.classList.add('correct');
            feedbackElement.textContent = 'சரியான பதில்!';
            feedbackElement.classList.add('correct');
        } else {
            option.classList.add('incorrect');
            document.querySelector(`.option[data-option="${question.key}"]`).classList.add('correct');
            feedbackElement.textContent = `தவறான பதில்! சரியான பதில்: ${question.key}`;
            feedbackElement.classList.add('incorrect');
        }
        
        updateSubjectScore(question.subject, isCorrect);
        clearInterval(timerInterval);
    });
});

// Button Listeners
nextButton.addEventListener('click', nextQuestion);
prevButton.addEventListener('click', prevQuestion);
restartButton.addEventListener('click', restartQuiz);
exitButton.addEventListener('click', exitQuiz);
loginBtn.addEventListener('click', loginUser);
logoutBtn.addEventListener('click', logoutUser);
applyFiltersBtn.addEventListener('click', applyFilters);
resetFiltersBtn.addEventListener('click', resetFilters);

userNameInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') loginUser();
});

// Start App
initApp();
