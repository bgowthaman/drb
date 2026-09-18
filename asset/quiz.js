// DOM elements
const questionElement = document.getElementById('question');
const optionElements = document.querySelectorAll('.option');
const feedbackElement = document.getElementById('feedback');
const prevButton = document.getElementById('prev-btn');
const nextButton = document.getElementById('next-btn');
const scoreElement = document.getElementById('score');
const totalQuestionsElement = document.getElementById('total-questions');
const progressBar = document.getElementById('progress-bar');
const excelFileInput = document.getElementById('excel-file');
const jsonFileInput = document.getElementById('json-file');
const endQuizElement = document.getElementById('end-quiz');
const finalScoreElement = document.getElementById('final-score');
const finalTotalElement = document.getElementById('final-total');
const restartButton = document.getElementById('restart-btn');
const exitButton = document.getElementById('exit-btn');
const fileInputContainer = document.getElementById('file-input-container');
const descriptionElement = document.getElementById('description');
const instructionsElement = document.getElementById('instructions');
const timerContainer = document.getElementById('timer-container');
const timerElement = document.getElementById('timer');
const quizContent = document.getElementById('quiz-content');
const scoreSummaryElement = document.getElementById('score-summary');

// HTML இல் உள்ள பட டேக்கைக் கண்டறிதல்
const imageElement = document.getElementById('questionImage');

// New filter elements
const filterContainer = document.getElementById('filter-container');
const yearButtonsContainer = document.getElementById('year-buttons');
const subjectButtonsContainer = document.getElementById('subject-buttons');
const typeButtonsContainer = document.getElementById('type-buttons'); // Added type container
const randomizeCheckbox = document.getElementById('randomize-checkbox');
const applyFiltersBtn = document.getElementById('apply-filters-btn');
const resetFiltersBtn = document.getElementById('reset-filters-btn');

// User elements
const userContainer = document.getElementById('user-container');
const welcomeMessage = document.getElementById('welcome-message');
const userNameInput = document.getElementById('user-name');
const loginBtn = document.getElementById('login-btn');
const logoutBtn = document.getElementById('logout-btn');
const loginSection = document.getElementById('login-section');
const userInfoSection = document.getElementById('user-info-section');

// Score elements
const scoreDisplay = document.getElementById('score-display');
const mainScoreBoard = document.getElementById('main-score-board');
const scoreDetails = document.getElementById('score-details');
const averageScoreElement = document.getElementById('average-score');

// Quiz state
let quizData = [];
let currentQuestionIndex = 0;
let score = 0;
let selectedOption = null;
let questionOrder = [];
let userAnswers = {}; // Store user answers and whether they were correct
let timerInterval = null;
let timeLeft = 30;		
let allQuestions = []; // Store all questions from JSON
let availableYears = [];
let availableSubjects = [];
let availableTypes = []; // Added available types
let selectedYears = [];
let selectedSubjects = [];
let selectedTypes = []; // Added selected types
let currentFilters = {};
let currentUserName = '';
let subjectWiseScores = {}; // Track scores by subject
let subjectWiseTotals = {}; // Track total questions by subject

// Auto-load question bank from default path
function autoLoadQuestionBank() {
    const defaultPath = './data/question_bank.xlsx';
    
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
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            
            // Process the data (skip header row)
            processExcelData(jsonData.slice(1));
        })
        .catch(error => {
            console.error('Auto-load error:', error);
            alert('வினா வங்கியை தானாக ஏற்ற முடியவில்லை. கோப்பு பாதை: ' + defaultPath);
        });
}

// Initialize the application
function initApp() {
    // Check if user exists in localStorage
    const savedUserName = localStorage.getItem('quizUserName');
    
    if (savedUserName) {
        // Welcome back the user
        currentUserName = savedUserName;
        welcomeMessage.textContent = `வணக்கம் ${savedUserName}!`;
        loginSection.classList.add('hidden');
        userInfoSection.classList.remove('hidden');
        
        // Show main score board
        showMainScoreBoard(savedUserName);
        
		// Show file input options
        //fileInputContainer.classList.remove('hidden');
        //descriptionElement.classList.remove('hidden');
        //instructionsElement.classList.remove('hidden');
		
		// Auto-load question bank instead of showing file input
		autoLoadQuestionBank();
	
    } else {
        // Show login form
        loginSection.classList.remove('hidden');
        userInfoSection.classList.add('hidden');
        
        // Hide file inputs until user logs in
        //fileInputContainer.classList.add('hidden');
        //descriptionElement.classList.add('hidden');
        //instructionsElement.classList.add('hidden');
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
        
        // Save username to localStorage
        localStorage.setItem('quizUserName', userName);
        
        // Show main score board
        showMainScoreBoard(userName);
        
        // Show file input options
        //fileInputContainer.classList.remove('hidden');
        //descriptionElement.classList.remove('hidden');
        //instructionsElement.classList.remove('hidden');
		
		// Auto-load question bank instead of showing file input
        autoLoadQuestionBank();
		
    } else {
        alert('தயவு செய்து உங்கள் பெயரை உள்ளிடவும்');
    }
}

// Show main score board with user's performance
function showMainScoreBoard(userName) {
    const userScores = JSON.parse(localStorage.getItem('userScores') || '{}');
    const userData = userScores[userName] || {};
    
    let scoresHTML = '';
    let totalScore = 0;
    let totalPossible = 0;
    let subjectAverages = {};
    let quizCounts = {}; // To track unique quizzes
    
    if (Object.keys(userData).length === 0) {
        scoresHTML = '<p>இதுவரை மதிப்பெண்கள் இல்லை</p>';
        averageScoreElement.textContent = '0';
    } else {
        // Calculate subject-wise averages
        subjectAverages = {};
        quizCounts = {};
        
        for (const [quizId, quizData] of Object.entries(userData)) {
            if (quizId === 'quizzes') continue;
            
            // Track unique quizzes
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
        
        // Calculate overall average
        const overallAverage = totalPossible > 0 ? Math.round((totalScore / totalPossible) * 100) : 0;
        averageScoreElement.textContent = overallAverage;
        
        // Add subject averages
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
        
        // Add overall quiz average
        const totalQuizzes = Object.keys(quizCounts).length;
        
        scoresHTML += `<div class="subject-score">
            <span><b>தேர்வுகள்: ${totalQuizzes}</b></span>
            <span><b>${overallAverage}% (${totalScore}/${totalPossible})</b></span>
        </div>`;
    }
    
    scoreDetails.innerHTML = scoresHTML;
    mainScoreBoard.style.display = 'block';
}

// Save user scores to localStorage with quiz-specific data
function saveUserScores() {
    const userScores = JSON.parse(localStorage.getItem('userScores') || '{}');
    
    if (!userScores[currentUserName]) {
        userScores[currentUserName] = { quizzes: {} };
    }
    
    // Create a unique quiz ID based on timestamp
    const quizId = `quiz_${Date.now()}`;
    
    // Store this quiz's data
    userScores[currentUserName][quizId] = {
        date: new Date().toISOString(),
        totalScore: score,
        totalQuestions: quizData.length,
        subjects: {...subjectWiseScores}
    };
    
    // Update the quizzes list
    if (!userScores[currentUserName].quizzes) {
        userScores[currentUserName].quizzes = {};
    }
    userScores[currentUserName].quizzes[quizId] = true;
    
    localStorage.setItem('userScores', JSON.stringify(userScores));
    localStorage.setItem('quizUserName', currentUserName);
}

// Logout user
function logoutUser() {
    localStorage.removeItem('quizUserName');
    currentUserName = '';
    welcomeMessage.textContent = 'வணக்கம்! உங்கள் பெயரை உள்ளிடவும்';
    loginSection.classList.remove('hidden');
    userInfoSection.classList.add('hidden');
    
    // Also hide file inputs when logging out
	if (fileInputContainer) fileInputContainer.classList.add('hidden');
	if (descriptionElement) descriptionElement.classList.add('hidden');
	if (instructionsElement) instructionsElement.classList.add('hidden');
    
    // Hide filter container
    filterContainer.style.display = 'none';
    
    // Clear timer
    if (timerInterval) {
        clearInterval(timerInterval);
        timerInterval = null;
    }
    
    // Reset quiz state
    quizData = [];
    currentQuestionIndex = 0;
    score = 0;
    selectedOption = null;
    userAnswers = {};
    
    // Update UI
    quizContent.style.display = 'none';
    endQuizElement.style.display = 'none';
       
    // Clear the username input
    userNameInput.value = '';
    
    // Reloads the page
    window.location.reload(); // Reloads the page
}

// Handle file upload (Excel or JSON)
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const fileExtension = file.name.split('.').pop().toLowerCase();
    
    if (fileExtension === 'xlsx' || fileExtension === 'xls') {
        handleExcelFile(file);
    } else if (fileExtension === 'json') {
        handleJsonFile(file);
    } else {
        alert('தவறான கோப்பு வடிவம். Excel (.xlsx, .xls) அல்லது JSON (.json) கோப்புகளை மட்டும் பயன்படுத்தவும்.');
    }
}

// Handle Excel file
function handleExcelFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Assuming the first sheet contains the quiz data
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        
        // Process the data (skip header row)
        processExcelData(jsonData.slice(1));
    };
    reader.readAsArrayBuffer(file);
}

// Process Excel data into quiz format
function processExcelData(data) {
    quizData = data.map(row => {
        // Process line breaks in Excel data
        const processText = (text) => {
            if (!text) return "";
            // First convert \n to <br>
            let processed = String(text).replace(/\n/g, '<br>');
            // Then handle any existing <br /> tags
            processed = processed.replace(/<br\s*\/?>/gi, '<br>');
            return processed;
        };
        
        // Adjust column indices based on your Excel structure
        return {
            year: processText(row[0]) || "வருடம்",
            subject: processText(row[1]) || "பாடம்",
            type: processText(row[2]) || "வடிவம்", // Added type field
            question: processText(row[4]) || "கேள்வி இல்லை",
            image_path: processText(row[10]) || " படம் இல்லை",
            options: {
                A: processText(row[5]) || "விருப்பம் A",
                B: processText(row[6]) || "விருப்பம் B",
                C: processText(row[7]) || "விருப்பம் C",
                D: processText(row[8]) || "விருப்பம் D"
            },
            key: String(row[9] || "A").charAt(0).toUpperCase(),					
        };
    }).filter(item => item.question !== "கேள்வி இல்லை"); // Fixed the filter condition
    
    if (quizData.length > 0) {
        // Store all questions
        allQuestions = quizData;
        
        // Extract available years, subjects, and types
        extractFilters(allQuestions);
        
        // Show filter options
        showFilterOptions();
    } else {
        questionElement.textContent = "கோப்பில் ஏற்றத்தக்க தரவு இல்லை";
    }
}

// Handle JSON file
function handleJsonFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const jsonData = JSON.parse(e.target.result);
            processJsonData(jsonData);
        } catch (error) {
            alert('JSON கோப்பைப் படிக்கும்போது பிழை: ' + error.message);
        }
    };
    reader.readAsText(file);
}

// Process JSON data into quiz format
function processJsonData(data) {
    // Check if data is an array
    if (!Array.isArray(data)) {
        alert('JSON கோப்பு ஒரு வரிசையைக் கொண்டிருக்க வேண்டும்');
        return;
    }
    
    // Process JSON data to ensure it has year, subject, and type fields with default values
    const processedData = data.map(item => {
        // Create a new object with default values
        const processedItem = {
            year: "வருடம்", // Default year
            subject: "பாடம்", // Default subject
            type: "வடிவம்", // Default type
            question: item.question || "கேள்வி இல்லை",
            image_path: item.image_path || " படம் இல்லை",
            options: {
                A: (item.options && item.options.A) || "விருப்பம் A",
                B: (item.options && item.options.B) || "விருப்பம் B",
                C: (item.options && item.options.C) || "விரு�ப்பம் C",
                D: (item.options && item.options.D) || "விருப்பம் D"
            },
            key: String(item.key || "A").charAt(0).toUpperCase()
        };
        
        // Override with actual values if they exist
        if (item.year) processedItem.year = item.year;
        if (item.subject) processedItem.subject = item.subject;
        if (item.type) processedItem.type = item.type; // Added type override
        
        return processedItem;
    }).filter(item => item.question !== "கேள்வி இல்லை");
    
    if (processedData.length === 0) {
        alert('JSON கோப்பில் ஏற்றத்தக்க தரவு இல்லை');
        return;
    }
    
    // Store all questions
    allQuestions = processedData;
        
    // Extract available years, subjects, and types
    extractFilters(allQuestions);
        
    // Show filter options
    showFilterOptions();
}

// Extract available years, subjects, and types from questions
function extractFilters(questions) {
    availableYears = ['all'];
    availableSubjects = ['all'];
    availableTypes = ['all']; // Initialize available types
    
    questions.forEach(item => {
        // Extract year if available
        if (item.year && !availableYears.includes(item.year)) {
            availableYears.push(item.year);
        }
        
        // Extract subject if available
        if (item.subject && !availableSubjects.includes(item.subject)) {
            availableSubjects.push(item.subject);
        }
        
        // Extract type if available
        if (item.type && !availableTypes.includes(item.type)) {
            availableTypes.push(item.type);
        }
    });
    
    // Create year buttons
    yearButtonsContainer.innerHTML = '';
    availableYears.forEach(year => {
        const button = document.createElement('button');
        button.className = 'filter-btn';
        button.textContent = year === 'all' ? 'அனைத்தும்' : year;
        button.setAttribute('data-value', year);
        button.addEventListener('click', () => toggleFilter('year', year, button));
        yearButtonsContainer.appendChild(button);
    });
    
    // Create subject buttons
    subjectButtonsContainer.innerHTML = '';
    availableSubjects.forEach(subject => {
        const button = document.createElement('button');
        button.className = 'filter-btn';
        button.textContent = subject === 'all' ? 'அனைத்தும்' : subject;
        button.setAttribute('data-value', subject);
        button.addEventListener('click', () => toggleFilter('subject', subject, button));
        subjectButtonsContainer.appendChild(button);
    });
    
    // Create type buttons
    typeButtonsContainer.innerHTML = '';
    availableTypes.forEach(type => {
        const button = document.createElement('button');
        button.className = 'filter-btn';
        button.textContent = type === 'all' ? 'அனைத்தும்' : type;
        button.setAttribute('data-value', type);
        button.addEventListener('click', () => toggleFilter('type', type, button));
        typeButtonsContainer.appendChild(button);
    });
    
    // Try to load saved filters from localStorage
    loadSavedFilters();
}

// Toggle filter selection
function toggleFilter(type, value, button) {
    if (value === 'all') {
        // Select all
        if (type === 'year') {
            selectedYears = [...availableYears];
            document.querySelectorAll('#year-buttons .filter-btn').forEach(btn => {
                btn.classList.add('selected');
            });
        } else if (type === 'subject') {
            selectedSubjects = [...availableSubjects];
            document.querySelectorAll('#subject-buttons .filter-btn').forEach(btn => {
                btn.classList.add('selected');
            });
        } else {
            selectedTypes = [...availableTypes];
            document.querySelectorAll('#type-buttons .filter-btn').forEach(btn => {
                btn.classList.add('selected');
            });
        }
        return;
    }
    
    // Toggle selection
    button.classList.toggle('selected');
    
    if (type === 'year') {
        if (button.classList.contains('selected')) {
            selectedYears.push(value);
        } else {
            selectedYears = selectedYears.filter(y => y !== value);
        }
        
        // Update "all" button state
        const allBtn = document.querySelector('#year-buttons .filter-btn[data-value="all"]');
        if (selectedYears.length === availableYears.length - 1) {
            allBtn.classList.add('selected');
            selectedYears.push('all');
        } else {
            allBtn.classList.remove('selected');
            selectedYears = selectedYears.filter(y => y !== 'all');
        }
    } else if (type === 'subject') {
        if (button.classList.contains('selected')) {
            selectedSubjects.push(value);
        } else {
            selectedSubjects = selectedSubjects.filter(s => s !== value);
        }
        
        // Update "all" button state
        const allBtn = document.querySelector('#subject-buttons .filter-btn[data-value="all"]');
        if (selectedSubjects.length === availableSubjects.length - 1) {
            allBtn.classList.add('selected');
            selectedSubjects.push('all');
        } else {
            allBtn.classList.remove('selected');
            selectedSubjects = selectedSubjects.filter(s => s !== 'all');
        }
    } else {
        // Handle type filter
        if (button.classList.contains('selected')) {
            selectedTypes.push(value);
        } else {
            selectedTypes = selectedTypes.filter(t => t !== value);
        }
        
        // Update "all" button state
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
    if (fileInputContainer) fileInputContainer.classList.add('hidden');
    if (descriptionElement) descriptionElement.classList.add('hidden');
    if (instructionsElement) instructionsElement.classList.add('hidden');
    
    // Show filter container
    filterContainer.style.display = 'block';
}                               // <-- stray }

// Apply filters and start quiz
function applyFilters() {
    const randomize = randomizeCheckbox.checked;
    
    // Save filters to localStorage
    saveFilters(selectedYears, selectedSubjects, selectedTypes, randomize);
    
    // Filter questions based on selections
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
        alert('தேர்ந்தெடுத்த வடிப்பான்களுக்கு ஏற்ற கேள்விகள் இல்லை. தயவு செய்து வேறு வடிப்பான்களைத் தேர்ந்தெடுக்கவும்.');
        return;
    }
    
    // Store current filters for restart
    currentFilters = {
        years: [...selectedYears],
        subjects: [...selectedSubjects],
        types: [...selectedTypes],
        randomize: randomize
    };
    
    // Process the filtered questions
    quizData = filteredQuestions.map(item => {
        // Process line breaks in question and options
        const processText = (text) => {
            if (!text) return "";
            // First convert \n to <br>
            let processed = String(text).replace(/\n/g, '<br>');
            // Then handle any existing <br /> tags
            processed = processed.replace(/<br\s*\/?>/gi, '<br>');
            return processed;
        };
        
        return {
            question: processText(item.question) || "கேள்வி இல்லை",
            image_path: item.image_path || " படம் இல்லை",
            options: {
                A: processText(item.options?.A) || "விருப்பம் A",
                B: processText(item.options?.B) || "விருப்பம் B",
                C: processText(item.options?.C) || "விருப்பம் C",
                D: processText(item.options?.D) || "விருப்பம் D"
            },
            key: String(item.answer || item.key || "A").charAt(0).toUpperCase(),
            year: item.year || "வருடம்",
            subject: item.subject || "பொது",
            type: item.type || "வடிவம்" // Added type
        };
    }).filter(item => item.question !== "கேள்வி இல்லை");
    
    if (quizData.length > 0) {
        // Hide filter container and start quiz
        filterContainer.style.display = 'none';
        startQuiz(randomize);
    } else {
        questionElement.textContent = "வடிகட்டலுக்குப் பிறகு ஏற்றத்தக்க தரவு இல்லை";
    }
}

// Save filters to localStorage
function saveFilters(years, subjects, types, randomize) {
    if (typeof(Storage) !== "undefined") {
        localStorage.setItem('quizFilters', JSON.stringify({
            years: years,
            subjects: subjects,
            types: types,
            randomize: randomize
        }));
    }
}

// Load saved filters from localStorage
function loadSavedFilters() {
    if (typeof(Storage) !== "undefined") {
        const savedFilters = localStorage.getItem('quizFilters');
        if (savedFilters) {
            const filters = JSON.parse(savedFilters);
            
            // Apply year filters
            if (filters.years) {
                selectedYears = filters.years;
                document.querySelectorAll('#year-buttons .filter-btn').forEach(btn => {
                    const value = btn.getAttribute('data-value');
                    if (selectedYears.includes(value)) {
                        btn.classList.add('selected');
                    }
                });
            }
            
            // Apply subject filters
            if (filters.subjects) {
                selectedSubjects = filters.subjects;
                document.querySelectorAll('#subject-buttons .filter-btn').forEach(btn => {
                    const value = btn.getAttribute('data-value');
                    if (selectedSubjects.includes(value)) {
                        btn.classList.add('selected');
                    }
                });
            }
            
            // Apply type filters
            if (filters.types) {
                selectedTypes = filters.types;
                document.querySelectorAll('#type-buttons .filter-btn').forEach(btn => {
                    const value = btn.getAttribute('data-value');
                    if (selectedTypes.includes(value)) {
                        btn.classList.add('selected');
                    }
                });
            }
            
            // Apply randomize setting
            if (filters.randomize !== undefined) {
                randomizeCheckbox.checked = filters.randomize;
            }
        } else {
            // Default to selecting all
            selectedYears = [...availableYears];
            selectedSubjects = [...availableSubjects];
            selectedTypes = [...availableTypes];
            document.querySelectorAll('.filter-btn').forEach(btn => {
                btn.classList.add('selected');
            });
        }
    }
}

// Reset filters to default
function resetFilters() {
    selectedYears = [...availableYears];
    selectedSubjects = [...availableSubjects];
    selectedTypes = [...availableTypes];
    
    document.querySelectorAll('.filter-btn').forEach(btn => {
        btn.classList.add('selected');
    });
    
    randomizeCheckbox.checked = true;
}

// Start the quiz
function startQuiz(randomize = true) {
    // Initialize subject-wise scores
    subjectWiseScores = {};
    subjectWiseTotals = {};			

    // Set up question order
    questionOrder = Array.from({length: quizData.length}, (_, i) => i);
    
    // Reset quiz state
    currentQuestionIndex = 0;
    score = 0;
    selectedOption = null;
    userAnswers = {};
    
    // Update UI
    scoreElement.textContent = score;
    totalQuestionsElement.textContent = quizData.length;
    progressBar.style.width = '0%';
    
    // Show quiz content
    quizContent.style.display = 'block';
    endQuizElement.style.display = 'none';
    if (fileInputContainer) fileInputContainer.classList.add('hidden');
	if (descriptionElement) descriptionElement.classList.add('hidden');
	if (instructionsElement) instructionsElement.classList.add('hidden');
    
    // Display first question
    displayQuestion();
}

// Add this function to handle image display
function displayQuestionImage() {
    const questionIndex = questionOrder[currentQuestionIndex];
    const question = quizData[questionIndex];
    
    // Check if the question has an image path
    if (question.image_path && question.image_path !== "" && question.image_path !== " படம் இல்லை") {
        // Add the base path to the image
        const baseImagePath = "./data/image/";
        const fullImagePath = baseImagePath + question.image_path;
        
        imageElement.src = fullImagePath; 
        imageElement.style.display = 'block';
        imageElement.alt = "கேள்விக்குரிய படம்";
        
        // Add error handling for broken images
        imageElement.onerror = function() {
            console.error("Failed to load image:", fullImagePath);
            imageElement.style.display = 'none';
        };
    } else {
        imageElement.style.display = 'none';
    }
}


// Display current question
function displayQuestion() {
    const questionIndex = questionOrder[currentQuestionIndex];
    const question = quizData[questionIndex];
    
    // Update question text
    questionElement.innerHTML = question.question;
    
    // Display question image
    displayQuestionImage();
    
    // Update options
    optionElements.forEach(option => {
        const optionKey = option.getAttribute('data-option');
        option.innerHTML = question.options[optionKey];
        
        // Reset classes
        option.className = 'option tamil-font';
        
        // If user already answered this question, show the result
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
                                
    // Update navigation buttons
    prevButton.disabled = currentQuestionIndex === 0;
    nextButton.disabled = currentQuestionIndex === quizData.length;
    nextButton.textContent = currentQuestionIndex === quizData.length - 1 ? 'முடிவு' : 'அடுத்த கேள்வி';
    
    // Update progress bar
    progressBar.style.width = `${((currentQuestionIndex + 1) / quizData.length) * 100}%`;
    
    // Clear feedback
    feedbackElement.className = 'feedback';
    feedbackElement.textContent = '';
    
    // Start timer
    startTimer();
}

// Start timer for current question
function startTimer() {
    // Clear any existing timer
    if (timerInterval) {
        clearInterval(timerInterval);
    }
    
    timeLeft = 30;
    timerElement.textContent = timeLeft;
    timerContainer.classList.remove('timer-warning');
    
    timerInterval = setInterval(() => {
        timeLeft--;
        timerElement.textContent = timeLeft;
        
        // Add warning class when time is running out
        if (timeLeft <= 10) {
            timerContainer.classList.add('timer-warning');
        }
        
        // Time's up
        if (timeLeft <= 0) {
            clearInterval(timerInterval);
            handleTimeUp();
        }
    }, 1000);
}

// Handle time up for current question
function handleTimeUp() {
    const questionIndex = questionOrder[currentQuestionIndex];
    const question = quizData[questionIndex];
    
    // Mark as incorrect if not answered
    if (!userAnswers[questionIndex]) {
        userAnswers[questionIndex] = {
            selected: null,
            correct: false,
            score: 0  // Explicitly store zero score for unanswered questions
        };

        // Show correct answer
        optionElements.forEach(option => {
            const optionKey = option.getAttribute('data-option');
            if (optionKey === (question.key || question.answer || question.correct)) {
                option.classList.add('correct');
            }
        });
        
        // Show feedback
        feedbackElement.className = 'feedback incorrect';
        feedbackElement.innerHTML = 'நேரம் முடிந்துவிட்டது! சரியான பதில்: ' + (question.key || question.answer || question.correct);
    }
    
    // Enable next button
    nextButton.disabled = false;
}

// Update subject-wise scores
function updateSubjectScore(subject, isCorrect) {
    if (!subjectWiseScores[subject]) {
        subjectWiseScores[subject] = { score: 0, total: 0 };
    }
    if (!subjectWiseTotals[subject]) {
        subjectWiseTotals[subject] = 0;
    }
    
    subjectWiseScores[subject].total++;
    subjectWiseTotals[subject]++;
    
    if (isCorrect) {
        subjectWiseScores[subject].score++;
    }
}

// Check answer and provide feedback
function checkAnswer(selectedOptionKey) {
    const questionIndex = questionOrder[currentQuestionIndex];
    const question = quizData[questionIndex];
    
    // Clear timer
    if (timerInterval) {
        clearInterval(timerInterval);
        timerInterval = null;
    }
    
    // Check if answer is correct - handle various answer formats
    const correctAnswer = question.key || question.answer || question.correct;
    const isCorrect = selectedOptionKey === correctAnswer;
    
    // Store user answer with explicit score value
    userAnswers[questionIndex] = {
        selected: selectedOptionKey,
        correct: isCorrect,
        score: isCorrect ? 1 : 0  // Explicitly store 1 for correct, 0 for incorrect
    };
    
    // Update score
    if (isCorrect) {
        score++;
    } else {
        // Explicitly ensure no score is added for incorrect answers
        // (though this is redundant as we're only incrementing for correct answers)
    }
    scoreElement.textContent = score;
    
    // Show feedback
    feedbackElement.className = isCorrect ? 'feedback correct' : 'feedback incorrect';
    feedbackElement.innerHTML = isCorrect ? 
    'சரியான பதில்! <i class="fas fa-check-circle"></i>' : 
    'தவறான பதில்! சரியான பதில்: ' + correctAnswer;
    
    // Highlight correct and incorrect options
    optionElements.forEach(option => {
        const optionKey = option.getAttribute('data-option');
        
        if (optionKey === correctAnswer) {
            option.classList.add('correct');
        } else if (optionKey === selectedOptionKey && !isCorrect) {
            option.classList.add('incorrect');
        }
        
        if (optionKey === selectedOptionKey) {
            option.classList.add('selected');
        }
    });
    
    // Enable next button
    nextButton.disabled = false;
}

// Move to next question
function nextQuestion() {
    if (currentQuestionIndex < quizData.length - 1) {
        currentQuestionIndex++;
        selectedOption = null;
        displayQuestion();
    } else {
        endQuiz();
    }
}

// Move to previous question
function prevQuestion() {
    if (currentQuestionIndex > 0) {
        currentQuestionIndex--;
        selectedOption = null;
        displayQuestion();
    }
}

// End the quiz
function endQuiz() {
    // Clear timer
    if (timerInterval) {
        clearInterval(timerInterval);
        timerInterval = null;
    }
    
    // Hide timer, question elements, and image
    timerContainer.style.display = 'none';
    questionElement.style.display = 'none';
    imageElement.style.display = 'none'; // Hide image at end of quiz
    document.querySelector('.options').style.display = 'none';
    feedbackElement.style.display = 'none';
    prevButton.style.display = 'none';
    nextButton.style.display = 'none';
    document.querySelector('.progress').style.display = 'none';
    document.querySelector('.score-container').style.display = 'none';
    
    // Show end quiz section
    endQuizElement.style.display = 'block';
    finalScoreElement.textContent = score;
    finalTotalElement.textContent = quizData.length;
    
    // Generate score summary
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
    
    // Save user scores
    saveUserScores();
}

// Restart the quiz
function restartQuiz() {
    // Reset UI elements
    timerContainer.style.display = 'block';
    questionElement.style.display = 'block';
    imageElement.style.display = 'block'; // Show image again on restart
    document.querySelector('.options').style.display = 'grid';
    feedbackElement.style.display = 'none';
    prevButton.style.display = 'block';
    nextButton.style.display = 'block';
    document.querySelector('.progress').style.display = 'block';
    document.querySelector('.score-container').style.display = 'block';
    endQuizElement.style.display = 'none';
    
    // If we have saved filters, apply them
    if (Object.keys(currentFilters).length > 0) {
        applySavedFilters();
    } else {
        // Otherwise start fresh
        startQuiz();
    }
}

// Exit quiz and return to file selection
function exitQuiz() {
    // Clear timer
    if (timerInterval) {
        clearInterval(timerInterval);
        timerInterval = null;
    }
    
    // Reset quiz state
    quizData = [];
    currentQuestionIndex = 0;
    score = 0;
    selectedOption = null;
    userAnswers = {};
    
    // Update UI
    quizContent.style.display = 'none';
    endQuizElement.style.display = 'none';
    
    // Show file input options again
    //fileInputContainer.classList.remove('hidden');
    //descriptionElement.classList.remove('hidden');
    //instructionsElement.classList.remove('hidden');
	    
    // Clear file inputs
    //excelFileInput.value = '';

    // Reload question bank automatically
    autoLoadQuestionBank();	
}

// Event listeners
optionElements.forEach(option => {
    option.addEventListener('click', () => {
        // If already answered this question, don't allow changing answer
        if (userAnswers[questionOrder[currentQuestionIndex]] !== undefined) {
            return;
        }
        
        // Remove previous selection
        optionElements.forEach(opt => opt.classList.remove('selected'));
        
        // Select this option
        option.classList.add('selected');
        selectedOption = option.getAttribute('data-option');
        
        // Check if answer is correct
        const question = quizData[questionOrder[currentQuestionIndex]];
        const isCorrect = selectedOption === question.key;
        
        // Record the answer
        userAnswers[questionOrder[currentQuestionIndex]] = {
            selected: selectedOption,
            correct: isCorrect
        };
        
        // Update score if correct
        if (isCorrect) {
            score++;
            scoreElement.textContent = score;
            option.classList.add('correct');
            feedbackElement.textContent = 'சரியான பதில்!';
            feedbackElement.classList.add('correct');
        } else {
            option.classList.add('incorrect');
            // Highlight correct answer
            document.querySelector(`.option[data-option="${question.key}"]`).classList.add('correct');
            feedbackElement.textContent = `தவறான பதில்! சரியான பதில்: ${question.key} (${question.options[question.key]})`;
            feedbackElement.classList.add('incorrect');
        }
        
        // Update subject-wise scores
        updateSubjectScore(question.subject, isCorrect);
        
        // Stop the timer
        clearInterval(timerInterval);
        
        // Auto-advance to next question after a delay if not the last question			
            
        if (currentQuestionIndex < quizData.length - 1) {
            setTimeout(() => {
                //nextQuestion();			
            }, 3000);
        } else {
            // If it's the last question, wait a bit then end the quiz
            setTimeout(() => {
                //endQuiz();
            }, 3000);
        }
    });
});

nextButton.addEventListener('click', nextQuestion);
prevButton.addEventListener('click', prevQuestion);
restartButton.addEventListener('click', restartQuiz);
exitButton.addEventListener('click', exitQuiz);

// User event listeners
loginBtn.addEventListener('click', loginUser);
logoutBtn.addEventListener('click', logoutUser);

// New filter event listeners
applyFiltersBtn.addEventListener('click', applyFilters);
resetFiltersBtn.addEventListener('click', resetFilters);

// Allow pressing Enter to login
userNameInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
        loginUser();
    }
});

// Initialize the app
initApp();