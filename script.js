// Hall Ticket Generator JavaScript

// Global configuration object
let hallTicketConfig = {
    schoolName: "GLOBAL'S SANMARG PUBLIC SCHOOL BIDAR",
    schoolSubtitle: "English Medium School With Shiksha-E-Hind and IIT Foundation Course",
    schoolAddress: "Main Branch: Pansal Talleem  Bidar           Branch: Near City Palace", 
    examTitle: "Marks Card Annual Exam (2020-21)",
    maxMarks: 100,
    minMarks: 35,
    subjects: {
        subject1: "Mathematics",
        subject2: "Science", 
        subject3: "Social",
        subject4: "English",
        subject5: "Kannada",
        subject6: "Hindi/Urdu"
    },
    // Auto-calculated properties
    get totalMaxMarks() {
        const configuredSubjectsCount = this.getConfiguredSubjectsCount();
        return this.maxMarks * configuredSubjectsCount;
    },
    get totalMinMarks() {
        const configuredSubjectsCount = this.getConfiguredSubjectsCount();
        return this.minMarks * configuredSubjectsCount;
    },
    // Helper method to count configured subjects
    getConfiguredSubjectsCount() {
        let count = 0;
        for (let i = 1; i <= 6; i++) {
            const subjectName = this.subjects[`subject${i}`];
            if (subjectName && subjectName.trim() !== '') {
                count++;
            }
        }
        return count > 0 ? count : 6; // Default to 6 if none configured
    }
};

// Global variables
let currentStep = 1;
let uploadedStudentsData = null;
let studentsData = [];

document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

function initializeApp() {
    console.log('Initializing Hall Ticket Generator...');
    
    // Check if XLSX library is loaded
    if (typeof XLSX !== 'undefined') {
        console.log('✅ XLSX library loaded successfully');
    } else {
        console.error('❌ XLSX library not loaded!');
    }
    
    // Check if html2pdf is loaded
    if (typeof html2pdf !== 'undefined') {
        console.log('✅ html2pdf library loaded successfully');
    } else {
        console.error('❌ html2pdf library not loaded!');
    }
    
    // Setup event listeners
    setupFileUpload();
    setupStepNavigation();
    setupConfigurationListeners();
    
    // Initialize first step
    showStep(1);
    updateProgressSteps(1);
    
    console.log('Hall Ticket Generator initialized successfully');
}

function setupFileUpload() {
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('excelFile');
    const browseBtn = document.getElementById('browseBtn');
    
    if (!uploadArea || !fileInput || !browseBtn) {
        console.error('Upload elements not found');
        return;
    }
    
    // Click to browse
    browseBtn.addEventListener('click', () => {
        fileInput.click();
    });
    
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });
    
    // File selection
    fileInput.addEventListener('change', handleFileSelect);
    
    // Drag and drop
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.style.borderColor = '#667eea';
        uploadArea.style.background = '#f8f9ff';
    });
    
    uploadArea.addEventListener('dragleave', (e) => {
        e.preventDefault();
        uploadArea.style.borderColor = '#ddd';
        uploadArea.style.background = '';
    });
    
    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.style.borderColor = '#ddd';
        uploadArea.style.background = '';
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            fileInput.files = files;
            handleFileSelect({ target: { files: files } });
        }
    });
}

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    console.log('File selected:', file.name);
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            console.log('Excel data loaded:', jsonData);
            
            // Process the data
            processExcelData(jsonData);
            
        } catch (error) {
            console.error('Error reading file:', error);
            alert('Error reading file. Please make sure it\'s a valid Excel file.');
        }
    };
    
    reader.readAsArrayBuffer(file);
}

function processExcelData(data) {
    if (data.length < 2) {
        alert('Excel file must have at least a header row and one data row.');
        return;
    }
    
    const headers = data[0].map(h => String(h).toLowerCase().replace(/[^a-z0-9]/g, ''));
    const rows = data.slice(1);
    
    console.log('Headers:', headers);
    console.log('Rows:', rows.length);
    
    studentsData = [];
    
    rows.forEach((row, index) => {
        if (row.length === 0) return; // Skip empty rows
        
        const student = {
            rollNumber: '',
            studentName: '',
            fatherName: '',
            class: '',
            section: '',
            marks: {
                mathematics: 0,
                science: 0,
                social: 0,
                english: 0,
                kannada: 0,
                hindi: 0
            }
        };
        
        // Map Excel columns to student data
        headers.forEach((header, colIndex) => {
            const value = row[colIndex] || '';
            
            // Map different possible column names
            if (header.includes('roll') || header.includes('number')) {
                student.rollNumber = value;
            } else if (header.includes('student') && header.includes('name')) {
                student.studentName = value;
            } else if (header.includes('father') && header.includes('name')) {
                student.fatherName = value;
            } else if (header.includes('class')) {
                student.class = value;
            } else if (header.includes('section')) {
                student.section = value;
            } else if (header.includes('math')) {
                student.marks.mathematics = value === '' || value === null || value === undefined ? 0 : parseInt(value) || 0;
            } else if (header.includes('science')) {
                student.marks.science = value === '' || value === null || value === undefined ? 0 : parseInt(value) || 0;
            } else if (header.includes('social')) {
                student.marks.social = value === '' || value === null || value === undefined ? 0 : parseInt(value) || 0;
            } else if (header.includes('english')) {
                student.marks.english = value === '' || value === null || value === undefined ? 0 : parseInt(value) || 0;
            } else if (header.includes('kannada')) {
                student.marks.kannada = value === '' || value === null || value === undefined ? 0 : parseInt(value) || 0;
            } else if (header.includes('hindi') || header.includes('urdu')) {
                student.marks.hindi = value === '' || value === null || value === undefined ? 0 : parseInt(value) || 0;
            }
        });
        
        // Only add students with required data
        if (student.rollNumber && student.studentName) {
            studentsData.push(student);
        }
    });
    
    console.log('Processed students:', studentsData.length);
    
    if (studentsData.length === 0) {
        alert('No valid student data found. Please check your Excel format.');
        return;
    }
    
    // Update UI
    uploadedStudentsData = studentsData;
    updateUploadStatus();
    enableNextButton();
}

function updateUploadStatus() {
    const uploadArea = document.getElementById('uploadArea');
    if (uploadArea && studentsData.length > 0) {
        uploadArea.innerHTML = `
            <div class="upload-icon">✅</div>
            <h3>File Uploaded Successfully!</h3>
            <p>Found ${studentsData.length} students</p>
            <p>Ready to proceed to customization</p>
        `;
        uploadArea.style.borderColor = '#28a745';
        uploadArea.style.background = '#f8fff9';
    }
}

function enableNextButton() {
    const nextBtn = document.getElementById('nextToCustomize');
    if (nextBtn) {
        nextBtn.disabled = false;
        nextBtn.classList.remove('btn-disabled');
    }
}

function setupStepNavigation() {
    // Next buttons
    const nextToCustomize = document.getElementById('nextToCustomize');
    const nextToMarks = document.getElementById('nextToMarks');
    const nextToPreview = document.getElementById('nextToPreview');
    
    if (nextToCustomize) nextToCustomize.addEventListener('click', () => goToStep(2));
    if (nextToMarks) nextToMarks.addEventListener('click', () => goToStep(3));
    if (nextToPreview) nextToPreview.addEventListener('click', () => goToStep(4));
    
    // Back buttons
    const backToUpload = document.getElementById('backToUpload');
    const backToCustomize = document.getElementById('backToCustomize');
    const backToMarks = document.getElementById('backToMarks');
    
    if (backToUpload) backToUpload.addEventListener('click', () => goToStep(1));
    if (backToCustomize) backToCustomize.addEventListener('click', () => goToStep(2));
    if (backToMarks) backToMarks.addEventListener('click', () => goToStep(3));
    
    // Preview and generation buttons
    const generatePreview = document.getElementById('generatePreview');
    const proceedToGenerate = document.getElementById('proceedToGenerate');
    
    if (generatePreview) generatePreview.addEventListener('click', generateHallTicketPreview);
    if (proceedToGenerate) proceedToGenerate.addEventListener('click', () => {
        goToStep(5);
        proceedToFinalGeneration();
    });
    
    // Make progress steps clickable
    document.querySelectorAll('.step').forEach(step => {
        step.addEventListener('click', () => {
            const stepNumber = parseInt(step.getAttribute('data-step'));
            if (stepNumber === 1 || uploadedStudentsData) {
                goToStep(stepNumber);
            } else {
                alert('Please upload an Excel file first!');
            }
        });
        step.style.cursor = 'pointer';
    });
}

function setupConfigurationListeners() {
    // Listen for changes in configuration inputs
    const configInputs = ['schoolName', 'schoolSubtitle', 'schoolAddress', 'examTitle', 'batchMaxMarks', 'batchMinMarks'];
    
    configInputs.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            element.addEventListener('input', updateConfiguration);
        }
    });
    
    // Subject configuration
    for (let i = 1; i <= 6; i++) {
        const element = document.getElementById(`subject${i}`);
        if (element) {
            element.addEventListener('input', updateConfiguration);
        }
    }
}

function updateConfiguration() {
    // Update global configuration from inputs
    const schoolName = document.getElementById('schoolName');
    const schoolSubtitle = document.getElementById('schoolSubtitle');
    const schoolAddress = document.getElementById('schoolAddress');
    const examTitle = document.getElementById('examTitle');
    const batchMaxMarks = document.getElementById('batchMaxMarks');
    const batchMinMarks = document.getElementById('batchMinMarks');
    
    if (schoolName) hallTicketConfig.schoolName = schoolName.value;
    if (schoolSubtitle) hallTicketConfig.schoolSubtitle = schoolSubtitle.value;
    if (schoolAddress) hallTicketConfig.schoolAddress = schoolAddress.value;
    if (examTitle) hallTicketConfig.examTitle = examTitle.value;
    if (batchMaxMarks) hallTicketConfig.maxMarks = parseInt(batchMaxMarks.value) || 100;
    if (batchMinMarks) hallTicketConfig.minMarks = parseInt(batchMinMarks.value) || 35;
    
    // Update subjects
    for (let i = 1; i <= 6; i++) {
        const element = document.getElementById(`subject${i}`);
        if (element) {
            hallTicketConfig.subjects[`subject${i}`] = element.value;
        }
    }
    
    // Update totals display
    updateBatchTotals();
}

function updateBatchTotals() {
    // Get current values from inputs
    const batchMaxMarks = document.getElementById('batchMaxMarks');
    const batchMinMarks = document.getElementById('batchMinMarks');
    
    if (batchMaxMarks && batchMinMarks) {
        const maxValue = parseInt(batchMaxMarks.value) || 100;
        const minValue = parseInt(batchMinMarks.value) || 35;
        
        // Update configuration
        hallTicketConfig.maxMarks = maxValue;
        hallTicketConfig.minMarks = minValue;
        
        // Update display elements
        const totalMaxElement = document.getElementById('batchTotalMax');
        const totalMinElement = document.getElementById('batchTotalMin');
        const maxFormulaElement = document.getElementById('maxFormula');
        const minFormulaElement = document.getElementById('minFormula');
        
        // Get configured subjects count
        const configuredCount = hallTicketConfig.getConfiguredSubjectsCount();
        
        if (totalMaxElement) {
            totalMaxElement.textContent = maxValue * configuredCount;
        }
        if (totalMinElement) {
            totalMinElement.textContent = minValue * configuredCount;
        }
        
        // Update formulas with current values
        if (maxFormulaElement) {
            maxFormulaElement.textContent = `(${maxValue} × ${configuredCount} subjects)`;
        }
        if (minFormulaElement) {
            minFormulaElement.textContent = `(${minValue} × ${configuredCount} subjects)`;
        }
        
        // Update subjects count display
        const subjectsInfo = document.querySelector('.subjects-info .subjects-count');
        if (subjectsInfo) {
            subjectsInfo.textContent = `📚 ${configuredCount} Subjects Total`;
        }
        
        console.log('Updated totals:', {
            maxPerSubject: maxValue,
            minPerSubject: minValue,
            totalMax: maxValue * 6,
            totalMin: minValue * 6
        });
    }
}

function goToStep(stepNumber) {
    // Hide all step contents
    document.querySelectorAll('.step-content').forEach(content => {
        content.classList.remove('active');
    });
    
    // Show target step content
    const targetStep = document.getElementById(`step${stepNumber}`);
    if (targetStep) {
        targetStep.classList.add('active');
    }
    
    // Update progress
    updateProgressSteps(stepNumber);
    updateProgressBar(stepNumber);
    
    currentStep = stepNumber;
    
    // Step-specific actions
    if (stepNumber === 3) {
        updateBatchTotals();
    }
}

function updateProgressSteps(activeStep) {
    document.querySelectorAll('.step').forEach(step => {
        const stepNum = parseInt(step.getAttribute('data-step'));
        step.classList.remove('active');
        
        if (stepNum === activeStep) {
            step.classList.add('active');
        } else if (uploadedStudentsData || stepNum === 1) {
            step.classList.add('accessible');
        }
    });
}

function updateProgressBar(activeStep) {
    const progressFill = document.querySelector('.progress-fill');
    if (progressFill) {
        const percentage = (activeStep / 5) * 100;
        progressFill.style.width = `${percentage}%`;
    }
}

function showStep(stepNumber) {
    // Hide all steps
    document.querySelectorAll('.step-content').forEach(content => {
        content.classList.remove('active');
    });
    
    // Show target step
    const targetStep = document.getElementById(`step${stepNumber}`);
    if (targetStep) {
        targetStep.classList.add('active');
    }
}

function generateHallTicketPreview() {
    if (!studentsData || studentsData.length === 0) {
        alert('No student data available. Please upload an Excel file first.');
        return;
    }
    
    console.log('Generating statistics for', studentsData.length, 'students');
    
    // Update student count
    const studentCount = document.getElementById('studentCount');
    if (studentCount) {
        studentCount.textContent = studentsData.length;
    }
    
    // Generate statistics dashboard
    generateStatisticsDashboard();
}

function populateClassFilter() {
    const classFilter = document.getElementById('classFilter');
    if (!classFilter || !studentsData) return;
    
    // Get unique classes
    const classes = [...new Set(studentsData.map(student => student.class).filter(c => c))];
    
    // Clear existing options (except "All Classes")
    classFilter.innerHTML = '<option value="">All Classes</option>';
    
    // Add class options
    classes.forEach(className => {
        const option = document.createElement('option');
        option.value = className;
        option.textContent = className;
        classFilter.appendChild(option);
    });
}

function filterReportCards() {
    const classFilter = document.getElementById('classFilter');
    const selectedClass = classFilter ? classFilter.value : '';
    
    const allHallTickets = document.querySelectorAll('.hall-ticket');
    
    allHallTickets.forEach((ticket, index) => {
        const student = studentsData[index];
        if (!student) return;
        
        if (selectedClass === '' || student.class === selectedClass) {
            ticket.style.display = 'block';
        } else {
            ticket.style.display = 'none';
        }
    });
    
    // Update visible count
    const visibleCount = selectedClass === '' ? studentsData.length : 
        studentsData.filter(s => s.class === selectedClass).length;
    
    const studentCount = document.getElementById('studentCount');
    if (studentCount) {
        studentCount.innerHTML = `<strong>${visibleCount}</strong> students ${selectedClass ? `(Class: ${selectedClass})` : ''}`;
    }
}

function generateHallTicketHTML(student, index) {
    // Get configured subjects (only those with names)
    const configuredSubjects = [];
    const subjectMarks = {
        subject1: student.marks.mathematics,
        subject2: student.marks.science,
        subject3: student.marks.social,
        subject4: student.marks.english,
        subject5: student.marks.kannada,
        subject6: student.marks.hindi
    };
    
    // Build array of configured subjects
    for (let i = 1; i <= 6; i++) {
        const subjectName = hallTicketConfig.subjects[`subject${i}`];
        if (subjectName && subjectName.trim() !== '') {
            configuredSubjects.push({
                name: subjectName,
                marks: subjectMarks[`subject${i}`],
                className: ['math', 'science', 'social', 'english', 'kannada', 'hindi'][i-1]
            });
        }
    }
    
    // Calculate totals based on configured subjects only
    const totalObtained = configuredSubjects.reduce((sum, subject) => sum + (subject.marks || 0), 0);
    const actualMaxMarks = configuredSubjects.length * hallTicketConfig.maxMarks;
    const actualMinMarks = configuredSubjects.length * hallTicketConfig.minMarks;
    
    const percentage = actualMaxMarks > 0 ? ((totalObtained / actualMaxMarks) * 100).toFixed(1) : '0.0';
    const grade = getGrade(percentage);
    
    // Generate subject rows dynamically
    const subjectRows = configuredSubjects.map(subject => `
        <tr>
            <td>${subject.name}</td>
            <td>${hallTicketConfig.maxMarks}</td>
            <td>${hallTicketConfig.minMarks}</td>
            <td class="${subject.className}-marks">${subject.marks || 0}</td>
            <td class="${subject.className}-remark">${getRemark(subject.marks || 0, hallTicketConfig.minMarks)}</td>
        </tr>
    `).join('');
    
    return `
        <div class="hall-ticket">
            <div class="diamond-border-frame">
                <!-- Watermark Background -->
                <div class="watermark-container">
                    <img src="images/logo.jpg" alt="School Logo" class="watermark-logo">
                </div>
                
                <!-- Header Section -->
                <div class="hall-ticket-header">
                    <div class="school-logo">
                        <img src="images/logo.jpg" alt="School Logo" class="logo-img" onerror="this.style.display='none'">
                    </div>
                    <div class="school-info">
                        <h1 class="school-name">${hallTicketConfig.schoolName}</h1>
                        <p class="school-subtitle">${hallTicketConfig.schoolSubtitle}</p>
                        <p class="school-address">${hallTicketConfig.schoolAddress}</p>
                        <h2 class="exam-title">${hallTicketConfig.examTitle}</h2>
                    </div>
                </div>

                <!-- Student Information -->
                <div class="student-info">
                    <div class="student-left">
                        <div class="student-field">
                            <span class="field-label">STUDENT NAME:</span>
                            <span class="field-value">${student.studentName || 'SYEDA WANIYA MAHAM'}</span>
                        </div>
                        <div class="student-field">
                            <span class="field-label">CLASS:</span>
                            <span class="field-value">${student.class || '5th'}</span>
                        </div>
                        <div class="student-field">
                            <span class="field-label">ROLL NO:</span>
                            <span class="field-value">${student.rollNumber || '16'}</span>
                        </div>
                    </div>
                    <div class="student-right">
                        <div class="student-field">
                            <span class="field-label">FATHER NAME:</span>
                            <span class="field-value">${student.fatherName || 'SYED ZAHIR'}</span>
                        </div>
                        <div class="student-field">
                            <span class="field-label">SECTION:</span>
                            <span class="field-value">${student.section || 'A'}</span>
                        </div>
                    </div>
                </div>

                <!-- Marks Table -->
                <div class="marks-section">
                    <table class="marks-table">
                        <thead>
                            <tr>
                                <th>Subjects</th>
                                <th>Max.Marks</th>
                                <th>Min.Marks</th>
                                <th>Marks Obt.</th>
                                <th>Remark</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${subjectRows}
                            <tr class="total-row">
                                <td><strong>Total</strong></td>
                                <td><strong>${actualMaxMarks}</strong></td>
                                <td><strong>${actualMinMarks}</strong></td>
                                <td class="total-marks"><strong>${totalObtained}</strong></td>
                                <td></td>
                            </tr>
                            <tr class="percentage-row">
                                <td><strong>Percentage</strong></td>
                                <td colspan="3" class="percentage">${percentage}</td>
                                <td></td>
                            </tr>
                            <tr>
                                <td><strong>Grade</strong></td>
                                <td colspan="3" class="grade">${grade}</td>
                                <td></td>
                            </tr>
                        </tbody>
                    </table>
                </div>

                <!-- Co-Scholastic Areas -->
                <div class="co-scholastic">
                    <h3>Co-Scholastic Areas</h3>
                    <div class="co-scholastic-grid">
                        <div class="co-item">
                            <span>1. Discipline in the classroom :</span>
                            <div class="grade-options">[ A / B / C ]</div>
                        </div>
                        <div class="co-item">
                            <span>2. Behavior / Conduct with teachers & classmates :</span>
                            <div class="grade-options">[ A / B / C ]</div>
                        </div>
                        <div class="co-item">
                            <span>3. Regularity & Neatness in doing HW/CW :</span>
                            <div class="grade-options">[ A / B / C ]</div>
                        </div>
                        <div class="co-item">
                            <span>4. Comes to School :</span>
                            <div class="school-timing-options">
                                <span class="timing-option">[ &nbsp; ] On time</span>
                                <span class="timing-option">[ &nbsp; ] Sometimes Late</span>
                                <span class="timing-option">[ &nbsp; ] Always Late</span>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- Footer Signatures -->
                <div class="signatures">
                    <div class="signature-item">
                        <div class="signature-line"></div>
                        <strong>Parent Sign</strong>
                    </div>
                    <div class="signature-item">
                        <div class="signature-line"></div>
                        <strong>Class Teacher</strong>
                    </div>
                    <div class="signature-item">
                        <div class="signature-line"></div>
                        <strong>Academic Head</strong>
                    </div>
                </div>
            </div>

            <!-- Download Section -->
            <div class="download-section">
                <button class="print-btn" onclick="printHallTicket(${index})">🖨️ Print</button>
                <button class="view-btn" onclick="viewHallTicket(${index})">👁️ View PDF</button>
                <button class="download-btn" onclick="downloadHallTicket(${index})">📥 Download PDF</button>
            </div>
        </div>
    `;
}

function getRemark(obtained, minimum) {
    const obtainedNum = parseInt(obtained) || 0;
    const minimumNum = parseInt(minimum) || 35;
    console.log(`Checking remark: obtained=${obtainedNum}, minimum=${minimumNum}, result=${obtainedNum >= minimumNum ? 'Pass' : 'Fail'}`);
    return obtainedNum >= minimumNum ? '' : 'Fail';
}

function getGrade(percentage) {
    const p = parseFloat(percentage);
    if (p >= 90) return 'A+';
    if (p >= 80) return 'A';
    if (p >= 70) return 'B+';
    if (p >= 60) return 'B';
    if (p >= 50) return 'C+';
    if (p >= 40) return 'C';
    if (p >= 35) return 'D';
    return 'F';
}

function proceedToFinalGeneration() {
    const finalOutput = document.getElementById('finalOutput');
    if (!finalOutput) return;
    
    finalOutput.innerHTML = '';
    
    // Generate all hall tickets
    studentsData.forEach((student, index) => {
        const hallTicketHtml = generateHallTicketHTML(student, index);
        finalOutput.innerHTML += hallTicketHtml;
    });
    
    // Add download all button
    finalOutput.innerHTML += `
        <div style="text-align: center; margin: 30px; padding: 20px; background: #f8f9fa; border-radius: 10px;">
            <h3>📋 All Hall Tickets Generated!</h3>
            <p>Total: ${studentsData.length} hall tickets</p>
            <button class="btn btn-success" onclick="downloadAllHallTickets()" style="margin: 10px;">📥 Download All PDFs</button>
            <button class="btn btn-primary" onclick="window.print()" style="margin: 10px;">🖨️ Print All</button>
        </div>
    `;
}

// PDF Generation Functions
function printHallTicket(index) {
    const hallTicket = document.querySelectorAll('.hall-ticket')[index];
    if (!hallTicket) {
        alert('Hall ticket not found.');
        return;
    }
    
    // Clone the hall ticket to avoid modifying the original
    const ticketClone = hallTicket.cloneNode(true);
    
    // Remove download buttons from clone
    const downloadSection = ticketClone.querySelector('.download-section');
    if (downloadSection) {
        downloadSection.remove();
    }
    
    // Create print window with exact styling
    const printWindow = window.open('', '_blank');
    const printContent = `
        <!DOCTYPE html>
        <html>
            <head>
                <title>Hall Ticket - Print</title>
                <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
                <meta http-equiv="Pragma" content="no-cache">
                <meta http-equiv="Expires" content="0">
                <style>
                    * {
                        margin: 0;
                        padding: 0;
                        box-sizing: border-box;
                    }
                    
                    body {
                        font-family: 'Times New Roman', Times, serif;
                        background: white;
                        padding: 20px;
                    }
                    
                    .hall-ticket {
                        max-width: 900px;
                        margin: 0 auto;
                        background: white;
                        font-family: 'Times New Roman', Times, serif;
                    }
                    
                    .diamond-border-frame {
                        position: relative;
                        background: white;
                        margin: 20px;
                        padding: 30px;
                        border: 3px solid #8B4513;
                        overflow: visible;
                        box-sizing: border-box;
                    }
                    
                    /* Watermark Container */
                    .watermark-container {
                        position: absolute;
                        top: 50%;
                        left: 50%;
                        transform: translate(-50%, -50%);
                        z-index: 1;
                        pointer-events: none;
                        width: 450px;
                        height: 450px;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                    }
                    
                    .watermark-logo {
                        width: 450px;
                        height: 450px;
                        object-fit: contain;
                        opacity: 0.08;
                        max-width: 100%;
                        max-height: 100%;
                        display: block;
                    }
                    
                    .hall-ticket-header {
                        display: flex;
                        align-items: center;
                        margin-bottom: 20px;
                        padding-bottom: 5px;
                        position: relative;
                        z-index: 2;
                    }
                    
                    .school-logo {
                        margin-right: 15px;
                    }
                    
                    .logo-img {
                        width: 100px;
                        height: 100px;
                        object-fit: contain;
                        padding: 5px;
                        margin-left: 35px;
                    }
                    
                    .school-info {
                        width: 600px;
                        text-align: center;
                        margin: 0 auto;
                        margin-left: -40px;
                    }
                    
                    .school-name {
                        font-size: 20px;
                        font-weight: bold;
                        color: #000080;
                        margin: 0;
                        text-transform: uppercase;
                    }
                    
                    .school-subtitle {
                        font-size: 13px;
                        font-weight: bold;
                        color: #000;
                        margin: 0;
                    }
                    
                    .school-address {
                        font-size: 13px;
                        font-weight: bold;
                        color: #000;
                        margin: 0 0 15px 0;
                        white-space: pre;
                    }
                    
                    .exam-title {
                        font-size: 14px;
                        font-weight: bold;
                        color: #d00;
                        margin: 0;
                        text-transform: uppercase;
                    }
                    
                    .student-info {
                        display: flex;
                        justify-content: space-between;
                        margin-bottom: 20px;
                        padding-bottom: 15px;
                        position: relative;
                        z-index: 2;
                    }
                    
                    .student-left, .student-right {
                        flex: 1;
                    }
                    
                    .student-left {
                        margin-right: 30px;
                    }
                    
                    .student-field {
                        margin: 8px 0;
                        font-size: 12px;
                        display: flex;
                        align-items: baseline;
                    }
                    
                    .field-label {
                        color: #d00;
                        font-weight: normal;
                        margin-right: 5px;
                        min-width: 100px;
                    }
                    
                    .field-value {
                        color: #000080;
                        font-weight: bold;
                        text-transform: uppercase;
                        font-family: cursive;
                    }
                    
                    .marks-section {
                        margin: 20px 0;
                        position: relative;
                        z-index: 2;
                    }
                    
                    .marks-table {
                        width: 100%;
                        border-collapse: collapse;
                        font-size: 11px;
                        margin: 0;
                    }
                    
                    .marks-table th,
                    .marks-table td {
                        border: 1px solid #000;
                        padding: 6px 8px;
                        text-align: center;
                    }
                    
                    .marks-table th {
                        background-color: #ffcccc;
                        font-weight: bold;
                        font-size: 10px;
                        -webkit-print-color-adjust: exact;
                        print-color-adjust: exact;
                    }
                    
                    .marks-table td:first-child {
                        text-align: left;
                        font-weight: bold;
                        background-color: #fafafa;
                    }
                    
                    .marks-table .math-marks,
                    .marks-table .science-marks,
                    .marks-table .social-marks,
                    .marks-table .english-marks,
                    .marks-table .kannada-marks,
                    .marks-table .hindi-marks,
                    .marks-table .total-marks {
                        font-weight: bold;
                    }
                    
                    /* Remark styling */
                    .marks-table .math-remark,
                    .marks-table .science-remark,
                    .marks-table .social-remark,
                    .marks-table .english-remark,
                    .marks-table .kannada-remark,
                    .marks-table .hindi-remark {
                        color: #ff0000;
                        font-weight: bold;
                        font-size: 10px;
                    }
                    
                    .total-row td {
                        background-color: #ffcccc;
                        font-weight: bold;
                        -webkit-print-color-adjust: exact;
                        print-color-adjust: exact;
                    }
                    
                    .total-row td:first-child {
                        background-color: #ffcccc;
                        -webkit-print-color-adjust: exact;
                        print-color-adjust: exact;
                    }
                    
                    .percentage-row td {
                        background-color: #e8e8ff;
                    }
                    
                    .percentage {
                        color: #000080;
                        font-weight: bold;
                    }
                    
                    .co-scholastic {
                        margin: 20px 0;
                        padding-top: 15px;
                        position: relative;
                        z-index: 2;
                    }
                    
                    .co-scholastic h3 {
                        text-align: center;
                        font-weight: bold;
                        font-size: 12px;
                        margin: 10px 0;
                        text-decoration: underline;
                    }
                    
                    .co-scholastic-grid {
                        margin: 15px 0;
                    }
                    
                    .co-item {
                        display: flex;
                        justify-content: space-between;
                        align-items: center;
                        margin: 6px 0;
                        font-size: 11px;
                    }
                    
                    .co-item span {
                        flex: 1;
                        margin-right: 10px;
                    }
                    
                    .grade-options {
                        min-width: 120px;
                        text-align: center;
                        padding: 2px 5px;
                    }
                    
                    .school-timing-options {
                        display: flex;
                        gap: 15px;
                        align-items: center;
                        font-size: 11px;
                        flex-wrap: wrap;
                    }
                    
                    .timing-option {
                        display: flex;
                        align-items: center;
                        gap: 3px;
                        white-space: nowrap;
                    }
                    
                    .signatures {
                        display: flex;
                        justify-content: space-between;
                        margin-top: 30px;
                        padding-top: 20px;
                        position: relative;
                        z-index: 2;
                    }
                    
                    .signature-item {
                        text-align: center;
                        flex: 1;
                        margin: 0 15px;
                    }
                    
                    .signature-line {
                        height: 40px;
                        margin-bottom: 8px;
                    }
                    
                    .signature-item strong {
                        font-size: 10px;
                        font-weight: bold;
                    }
                    
                    @media print {
                        body {
                            padding: 0;
                            margin: 0;
                        }
                        
                        .hall-ticket {
                            margin: 0;
                            max-width: none;
                            width: 100%;
                        }
                        
                        @page {
                            margin: 0.5in;
                        }
                    }
                </style>
            </head>
            <body>
                ${ticketClone.outerHTML}
                <script>
                    window.onload = function() {
                        window.print();
                        window.onafterprint = function() {
                            window.close();
                        };
                    };
                </script>
            </body>
        </html>
    `;
    
    printWindow.document.write(printContent);
    printWindow.document.close();
}

function viewHallTicket(index) {
    const hallTicket = document.querySelectorAll('.hall-ticket')[index];
    if (!hallTicket) {
        alert('Hall ticket not found.');
        return;
    }
    
    const student = studentsData[index];
    const studentName = student.studentName || 'Student';
    
    // Hide download buttons temporarily
    const downloadSection = hallTicket.querySelector('.download-section');
    downloadSection.style.display = 'none';
    
    // Configure PDF options to match print dimensions exactly
    const opt = {
        margin: [12.7, 12.7, 12.7, 12.7], // 0.5in margins like print
        filename: `HallTicket_${studentName}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { 
            scale: 1.5,  // Increased for better quality
            useCORS: true,
            allowTaint: true,
            logging: false,
            letterRendering: true,
            backgroundColor: '#ffffff',
            scrollX: 0,
            scrollY: 0,
            width: 900,   // Match hall-ticket max-width from CSS
            height: 1200, // Proper height for full content
            x: 0,
            y: 0
        },
        jsPDF: { 
            unit: 'mm', 
            format: 'a4', 
            orientation: 'portrait',
            putOnlyUsedFonts: true,
            floatPrecision: 16
        }
    };
    
    // Generate PDF and open in new tab for viewing
    html2pdf().set(opt).from(hallTicket).toPdf().get('pdf').then(function (pdf) {
        const totalPages = pdf.internal.getNumberOfPages();
        for (let i = 1; i <= totalPages; i++) {
            pdf.setPage(i);
        }
        
        const blob = pdf.output('blob');
        const url = URL.createObjectURL(blob);
        const newTab = window.open(url, '_blank');
        newTab.document.title = `Hall Ticket - ${studentName}`;
        
        downloadSection.style.display = 'flex';
    }).catch((error) => {
        console.error('Error generating PDF preview:', error);
        alert('Error generating PDF preview. Please try downloading instead.');
        downloadSection.style.display = 'flex';
    });
}

function downloadHallTicket(index) {
    const hallTicket = document.querySelectorAll('.hall-ticket')[index];
    if (!hallTicket) {
        alert('Hall ticket not found.');
        return;
    }
    
    const student = studentsData[index];
    const studentName = student.studentName || 'Student';
    const rollNumber = student.rollNumber || 'RollNo';
    
    // Hide download buttons temporarily
    const downloadSection = hallTicket.querySelector('.download-section');
    downloadSection.style.display = 'none';
    
    // Configure PDF options to match print dimensions exactly
    const opt = {
        margin: [12.7, 12.7, 12.7, 12.7], // 0.5in margins like print
        filename: `HallTicket_${studentName.replace(/\s+/g, '_')}_${rollNumber}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { 
            scale: 1.5,  // Increased for better quality
            useCORS: true,
            allowTaint: true,
            logging: false,
            letterRendering: true,
            backgroundColor: '#ffffff',
            scrollX: 0,
            scrollY: 0,
            width: 900,   // Match hall-ticket max-width from CSS
            height: 1200, // Proper height for full content
            x: 0,
            y: 0
        },
        jsPDF: { 
            unit: 'mm', 
            format: 'a4', 
            orientation: 'portrait',
            putOnlyUsedFonts: true,
            floatPrecision: 16
        }
    };
    
    // Generate PDF and download
    html2pdf().set(opt).from(hallTicket).save().then(() => {
        downloadSection.style.display = 'flex';
    }).catch((error) => {
        console.error('Error generating PDF:', error);
        alert('Error generating PDF. Please try again.');
        downloadSection.style.display = 'flex';
    });
}

function addHallTicketContentToPDF(pdf, student) {
    const pageWidth = 210;
    const pageHeight = 297;
    const margin = 15;
    const startX = margin;
    const startY = margin;
    const cardWidth = pageWidth - (margin * 2);
    
    // Draw main border
    pdf.setDrawColor(102, 102, 102);
    pdf.setLineWidth(2);
    pdf.rect(startX, startY, cardWidth, pageHeight - (margin * 2));
    
    let currentY = startY + 15;
    
    // School Header
    pdf.setTextColor(0, 0, 0);
    pdf.setFontSize(16);
    pdf.setFont('times', 'bold');
    pdf.text(hallTicketConfig.schoolName, pageWidth/2, currentY, { align: 'center' });
    
    currentY += 8;
    pdf.setFontSize(11);
    pdf.setFont('times', 'bold');
    pdf.text(hallTicketConfig.schoolSubtitle, pageWidth/2, currentY, { align: 'center' });
    
    currentY += 8;
    pdf.setFontSize(12);
    pdf.setFont('times', 'bold');
    pdf.text(hallTicketConfig.examTitle, pageWidth/2, currentY, { align: 'center' });
    
    currentY += 15;
    
    // Horizontal line
    pdf.setLineWidth(0.5);
    pdf.line(startX + 5, currentY, startX + cardWidth - 5, currentY);
    currentY += 10;
    
    // Student Information
    pdf.setFontSize(10);
    pdf.setFont('times', 'normal');
    pdf.setTextColor(220, 0, 0); // Red color
    
    const leftX = startX + 10;
    const rightX = pageWidth/2 + 20;
    
    pdf.text(`STUDENT NAME: ${student.studentName || 'SYEDA WANIYA MAHAM'}`, leftX, currentY);
    pdf.text(`FATHER NAME: ${student.fatherName || 'SYED ZAHIR'}`, rightX, currentY);
    
    currentY += 8;
    pdf.text(`CLASS: ${student.class || '5th'}`, leftX, currentY);
    pdf.text(`SECTION: ${student.section || 'A'}`, rightX, currentY);
    
    currentY += 8;
    pdf.text(`ROLL NO: ${student.rollNumber || '16'}`, leftX, currentY);
    
    currentY += 15;
    
    // Marks Table
    pdf.setTextColor(0, 0, 0);
    addMarksTableToPDF(pdf, student, startX + 10, currentY, cardWidth - 20);
    
    currentY += 180; // Approximate table height
    
    // Co-Scholastic Areas
    addCoScholasticToPDF(pdf, startX + 10, currentY, cardWidth - 20);
    
    currentY += 50;
    
    // Signatures
    addSignaturesToPDF(pdf, startX + 10, currentY, cardWidth - 20);
}



function addMarksTableToPDF(pdf, student, x, y, width) {
    const rowHeight = 8;
    const colWidths = [50, 25, 25, 25, 25];
    let currentY = y;
    
    // Table header
    pdf.setFillColor(245, 245, 245);
    pdf.rect(x, currentY, width, rowHeight, 'F');
    pdf.setDrawColor(0, 0, 0);
    pdf.rect(x, currentY, width, rowHeight);
    
    pdf.setFontSize(9);
    pdf.setFont('times', 'bold');
    
    let currentX = x + 2;
    const headers = ['Subjects', 'Max.Marks', 'Min.Marks', 'Marks Obt.', 'Remark'];
    headers.forEach((header, i) => {
        pdf.text(header, currentX, currentY + 5);
        currentX += colWidths[i];
        if (i < headers.length - 1) {
            pdf.line(currentX, currentY, currentX, currentY + rowHeight);
        }
    });
    
    currentY += rowHeight;
    
    // Subject rows - only configured subjects
    const subjectMarks = {
        subject1: student.marks.mathematics,
        subject2: student.marks.science,
        subject3: student.marks.social,
        subject4: student.marks.english,
        subject5: student.marks.kannada,
        subject6: student.marks.hindi
    };
    
    const subjects = [];
    for (let i = 1; i <= 6; i++) {
        const subjectName = hallTicketConfig.subjects[`subject${i}`];
        if (subjectName && subjectName.trim() !== '') {
            subjects.push({
                name: subjectName,
                marks: subjectMarks[`subject${i}`] || 0
            });
        }
    }
    
    pdf.setFont('times', 'normal');
    
    subjects.forEach(subject => {
        pdf.rect(x, currentY, width, rowHeight);
        
        currentX = x + 2;
        const rowData = [
            subject.name,
            hallTicketConfig.maxMarks.toString(),
            hallTicketConfig.minMarks.toString(),
            subject.marks.toString(),
            getRemark(subject.marks, hallTicketConfig.minMarks)
        ];
        
        rowData.forEach((data, i) => {
            pdf.text(data, currentX, currentY + 5);
            currentX += colWidths[i];
            if (i < rowData.length - 1) {
                pdf.line(currentX, currentY, currentX, currentY + rowHeight);
            }
        });
        
        currentY += rowHeight;
    });
    
    // Total row
    const totalMarks = subjects.reduce((sum, s) => sum + s.marks, 0);
    const actualMaxMarks = subjects.length * hallTicketConfig.maxMarks;
    const actualMinMarks = subjects.length * hallTicketConfig.minMarks;
    const percentage = actualMaxMarks > 0 ? ((totalMarks / actualMaxMarks) * 100).toFixed(1) : '0.0';
    const grade = getGrade(percentage);
    
    pdf.setFillColor(240, 240, 240);
    pdf.rect(x, currentY, width, rowHeight, 'F');
    pdf.rect(x, currentY, width, rowHeight);
    
    pdf.setFont('times', 'bold');
    currentX = x + 2;
    const totalData = ['Total', actualMaxMarks.toString(), actualMinMarks.toString(), totalMarks.toString(), ''];
    totalData.forEach((data, i) => {
        pdf.text(data, currentX, currentY + 5);
        currentX += colWidths[i];
        if (i < totalData.length - 1) {
            pdf.line(currentX, currentY, currentX, currentY + rowHeight);
        }
    });
    
    currentY += rowHeight;
    
    // Percentage and Grade rows
    ['Rank', 'Percentage', 'Grade'].forEach((label, index) => {
        const value = index === 1 ? `${percentage}%` : index === 2 ? grade : '';
        
        if (index === 1) pdf.setFillColor(232, 232, 255);
        else if (index === 2) pdf.setFillColor(255, 245, 245);
        else pdf.setFillColor(255, 255, 255);
        
        pdf.rect(x, currentY, width, rowHeight, 'F');
        pdf.rect(x, currentY, width, rowHeight);
        
        pdf.text(label, x + 2, currentY + 5);
        if (value) pdf.text(value, x + colWidths[0] + 2, currentY + 5);
        
        pdf.line(x + colWidths[0], currentY, x + colWidths[0], currentY + rowHeight);
        
        currentY += rowHeight;
    });
}

function addCoScholasticToPDF(pdf, x, y, width) {
    pdf.setFont('times', 'bold');
    pdf.setFontSize(10);
    pdf.text('Co-Scholastic Areas', x + width/2, y, { align: 'center' });
    
    pdf.setFont('times', 'normal');
    pdf.setFontSize(8);
    
    const items = [
        '1. Discipline in the classroom',
        '2. Behavior / Conduct with teachers & classmates',
        '3. Regularity & Neatness in doing HW/CW',
        '4. Comes to School: On time [ ] / Sometimes Late [ ] / Always Late [ ]'
    ];
    
    let currentY = y + 8;
    items.forEach((item, index) => {
        pdf.text(item, x, currentY);
        if (index < 3) {
            pdf.text('[ A / B / C ]', x + width - 30, currentY);
        }
        currentY += 6;
    });
}

function addSignaturesToPDF(pdf, x, y, width) {
    pdf.setFont('times', 'bold');
    pdf.setFontSize(9);
    
    const signatures = ['Parent Sign', 'Class Teacher', 'Academic Head'];
    const sigWidth = width / 3;
    
    signatures.forEach((sig, index) => {
        const sigX = x + (index * sigWidth) + 20;
        pdf.text(sig, sigX + 20, y, { align: 'center' });
    });
}

function downloadAllHallTickets() {
    if (!studentsData || studentsData.length === 0) {
        alert('No student data available.');
        return;
    }
    
    try {
        const allHallTickets = document.querySelectorAll('.hall-ticket');
        if (allHallTickets.length === 0) {
            alert('No hall tickets found. Please generate them first.');
            return;
        }
        
        // Hide all download sections temporarily
        const downloadSections = document.querySelectorAll('.download-section');
        downloadSections.forEach(section => {
            section.style.display = 'none';
        });
        
        // Create a container with all hall tickets
        const container = document.createElement('div');
        container.style.width = '100%';
        container.style.backgroundColor = '#ffffff';
        
        allHallTickets.forEach((ticket, index) => {
            const ticketClone = ticket.cloneNode(true);
            // Remove download section from clone
            const downloadSection = ticketClone.querySelector('.download-section');
            if (downloadSection) {
                downloadSection.remove();
            }
            // Add page break after each ticket except the last
            if (index < allHallTickets.length - 1) {
                ticketClone.style.pageBreakAfter = 'always';
                ticketClone.style.marginBottom = '20px';
            }
            container.appendChild(ticketClone);
        });
        
        // Add container to body temporarily
        document.body.appendChild(container);
        
        const opt = {
            margin: [12.7, 12.7, 12.7, 12.7], // 0.5in margins like print
            filename: 'All_Hall_Tickets.pdf',
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { 
                scale: 1.5,  // Increased for better quality
                useCORS: true,
                allowTaint: true,
                logging: false,
                letterRendering: true,
                backgroundColor: '#ffffff',
                scrollX: 0,
                scrollY: 0,
                width: 900,   // Match hall-ticket max-width from CSS
                height: 1200, // Proper height for full content
                x: 0,
                y: 0
            },
            jsPDF: { 
                unit: 'mm', 
                format: 'a4', 
                orientation: 'portrait',
                putOnlyUsedFonts: true,
                floatPrecision: 16
            },
            pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
        };
        
        html2pdf().set(opt).from(container).save().then(() => {
            // Clean up
            document.body.removeChild(container);
            downloadSections.forEach(section => {
                section.style.display = 'flex';
            });
            alert(`Successfully generated ${studentsData.length} hall tickets!`);
        }).catch((error) => {
            console.error('Error generating all PDFs:', error);
            // Clean up
            if (document.body.contains(container)) {
                document.body.removeChild(container);
            }
            downloadSections.forEach(section => {
                section.style.display = 'flex';
            });
            alert('Error generating PDFs. Please try again.');
        });
        
    } catch (error) {
        console.error('Error generating all PDFs:', error);
        alert('Error generating PDFs. Please try again.');
    }
}

// Debug functions for testing
window.simpleTest = function() {
    console.log('🧪 Running simple test...');
    console.log('XLSX available:', typeof XLSX !== 'undefined');
    console.log('html2pdf available:', typeof html2pdf !== 'undefined');
    
    // Test with sample data including failing marks
    const sampleData = [
        ['ROLL_NUMBER', 'STUDENT_NAME', 'FATHER_NAME', 'CLASS', 'SECTION', 'MATHS', 'SCIENCE', 'SOCIAL', 'ENGLISH', 'KANNADA', 'HINDI'],
        [16, 'SYEDA WANIYA MAHAM', 'SYED ZAHIR', '5th', 'A', 22, 21, 5, 20, 21, 21],
        [1, 'RIYANSH', 'RAKESH', 'NURSERY', 'A', 25, 10, 2, 20, 21, 21]
    ];
    
    processExcelData(sampleData);
};

window.triggerFileInput = function() {
    const fileInput = document.getElementById('excelFile');
    if (fileInput) {
        fileInput.click();
        console.log('File input triggered');
    } else {
        console.log('File input not found');
    }
};