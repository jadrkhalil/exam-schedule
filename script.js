class ActionToastService {
    constructor() {
        this.toastElement = document.getElementById('toast-notification-right');
        this.messageElement = document.getElementById('toast-message-right');
        this.animationDuration = 3000;
    }

    show(message) {
        this.messageElement.textContent = message;
        this.toastElement.classList.remove('hidden');

        setTimeout(() => {
            this.toastElement.classList.add('hidden');
        }, this.animationDuration);
    }
}

// ==========================================
// FILE MANAGER SERVICE
// Handles all file upload operations
// ==========================================
class FileManager {
    constructor(fileInput, removeBtn, fileNameDisplay) {
        this.fileInput = fileInput;
        this.removeBtn = removeBtn;
        this.fileNameDisplay = fileNameDisplay;
        this.data = [];
    }

    initialize(onFileLoaded) {
        console.log('FileManager initializing...');
        console.log('File input element:', this.fileInput);
        console.log('File name display element:', this.fileNameDisplay);
        
        if (!this.fileInput) {
            console.error('File input element not found!');
            return;
        }
        
        this.fileInput.addEventListener('change', (e) => {
            console.log('File change event triggered');
            this.handleFileChange(e, onFileLoaded);
        });
        this.removeBtn.addEventListener('click', () => this.handleFileRemove());
    }

    handleFileChange(event, callback) {
        const file = event.target.files[0];
        console.log('handleFileChange called, file:', file);
        
        if (!file) {
            console.warn('No file selected');
            return;
        }

        this.fileNameDisplay.textContent = `✓ Loaded: ${file.name}`;
        this.removeBtn.classList.remove('hidden');

        const reader = new FileReader();
        reader.onload = (evt) => this.processFile(evt, callback);
        reader.readAsArrayBuffer(file);
    }

    processFile(event, callback) {
        console.log('processFile called');
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            console.log('Parsed rows:', jsonRows.length);
            this.data = this.parseExcelData(jsonRows);
            console.log('Parsed data:', this.data.length, 'courses');
            callback(this.data);
        } catch (error) {
            console.error('Error processing file:', error);
            this.fileNameDisplay.textContent = 'Error loading file. Please try again.';
        }
    }

    parseExcelData(rows) {
        const parsedData = [];
        let currentDate = 'Unknown Date';
        let currentTime = 'Unknown Time';

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;

            // Check if this is a date header
            if (row[0] && (!row[1] || row[1].toString().trim() === '')) {
                const text = row[0].toString().trim();
                if (text.match(/(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|202\d)/i)) {
                    currentDate = text;
                    currentTime = 'Unknown Time';
                }
                continue;
            }

            // Update current time
            if (row[0] && row[0].toString().trim() !== 'Time') {
                currentTime = row[0].toString().trim();
            }

            // Parse course data
            if (row[1] && row[1].toString().trim() !== 'Course #') {
                const courseCode = row[1].toString().trim();
                const normalized = this.normalizeCode(courseCode);
                const timestamp = this.parseTimestamp(currentDate, currentTime);

                parsedData.push({
                    dateStr: currentDate,
                    timeStr: currentTime,
                    timestamp: timestamp,
                    courseCode: courseCode,
                    normalized: normalized,
                    section: row[2] || '',
                    title: row[3] || '',
                    instructor: row[5] || '',
                    room: row[7] || ''
                });
            }
        }

        return parsedData;
    }

    normalizeCode(code) {
        if (!code) return '';
        return code.toString().replace(/\s+/g, '').toLowerCase();
    }

    parseTimestamp(dateStr, timeStr) {
        try {
            const startTime = timeStr.split('-')[0].trim();
            const cleanDate = dateStr.replace(/^[a-zA-Z]+,\s*/, '');
            return new Date(`${cleanDate} ${startTime}`).getTime();
        } catch (err) {
            console.warn('Could not parse timestamp:', dateStr, timeStr);
            return 0;
        }
    }

    handleFileRemove() {
        this.fileInput.value = '';
        this.data = [];
        this.fileNameDisplay.textContent = '✗ No file selected.';
        this.removeBtn.classList.add('hidden');
    }

    getLoadedData() {
        return this.data;
    }
}

// ==========================================
// SEARCH SERVICE
// Handles course search operations
// ==========================================
class SearchService {
    constructor(courseData) {
        this.courseData = courseData;
    }

    search(searchQuery) {
        const normalized = this.normalizeCode(searchQuery);
        if (!normalized) return [];
        return this.courseData.filter(item => item.normalized === normalized);
    }

    normalizeCode(code) {
        if (!code) return '';
        return code.toString().replace(/\s+/g, '').toLowerCase();
    }
}

// ==========================================
// SCHEDULE MANAGER
// Handles user schedule operations
// ==========================================
class ScheduleManager {
    constructor() {
        this.schedule = [];
    }

    addCourse(course) {
        const isDuplicate = this.schedule.some(
            s => s.normalized === course.normalized && s.section === course.section
        );

        if (isDuplicate) {
            return { success: false, message: `Section ${course.section} is already in your schedule!` };
        }

        this.schedule.push({ ...course });
        return { success: true };
    }

    removeCourse(index) {
        this.schedule.splice(index, 1);
    }

    clearSchedule() {
        this.schedule = [];
    }

    getSchedule() {
        return [...this.schedule];
    }

    getScheduleSorted() {
        return this.schedule.sort((a, b) => a.timestamp - b.timestamp);
    }
}

// ==========================================
// UI RENDERER
// Handles all DOM updates
// ==========================================
class UIRenderer {
    constructor() {
        this.courseSection = document.getElementById('course-section');
        this.scheduleSection = document.getElementById('schedule-section');
        this.searchResultsContainer = document.getElementById('search-results-container');
        this.resultsTbody = document.querySelector('#results-table tbody');
        this.scheduleTbody = document.querySelector('#schedule-table tbody');
        this.courseInput = document.getElementById('course-input');
    }

    showCourseSection() {
        this.courseSection.classList.remove('hidden');
        this.courseInput.focus();
    }

    hideCourseSection() {
        this.courseSection.classList.add('hidden');
    }

    showScheduleSection() {
        this.scheduleSection.classList.remove('hidden');
    }

    hideScheduleSection() {
        this.scheduleSection.classList.add('hidden');
    }

    renderSearchResults(results, onAddCourse) {
        this.resultsTbody.innerHTML = '';

        if (results.length === 0) {
            this.searchResultsContainer.classList.add('hidden');
            return false;
        }

        this.searchResultsContainer.classList.remove('hidden');

        results.forEach((item, index) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><span class="course-badge">${this.escapeHtml(item.courseCode)}</span></td>
                <td><strong>${this.escapeHtml(item.section)}</strong></td>
                <td>${this.escapeHtml(item.instructor)}</td>
                <td>
                    ${this.escapeHtml(item.dateStr)}
                    <span class="time-subtext">${this.escapeHtml(item.timeStr)}</span>
                </td>
                <td><button class="btn btn-success" data-index="${index}">Add</button></td>
            `;
            const btn = tr.querySelector('.btn-success');
            btn.addEventListener('click', () => onAddCourse(index));
            this.resultsTbody.appendChild(tr);
        });

        return true;
    }

    renderSchedule(schedule, onRemoveCourse) {
        this.scheduleTbody.innerHTML = '';

        if (schedule.length === 0) {
            this.hideScheduleSection();
            return;
        }

        this.showScheduleSection();

        schedule.forEach((item, index) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${this.escapeHtml(item.dateStr)}</td>
                <td>${this.escapeHtml(item.timeStr)}</td>
                <td><span class="course-badge">${this.escapeHtml(item.courseCode)}</span></td>
                <td>${this.escapeHtml(item.section)}</td>
                <td>${this.escapeHtml(item.title)}</td>
                <td>${this.escapeHtml(item.instructor)}</td>
                <td>${this.escapeHtml(item.room)}</td>
                <td><button class="remove-row-btn" data-index="${index}">✕ Remove</button></td>
            `;
            const btn = tr.querySelector('.remove-row-btn');
            btn.addEventListener('click', () => onRemoveCourse(index));
            this.scheduleTbody.appendChild(tr);
        });
    }

    clearSearchInput() {
        this.courseInput.value = '';
    }

    hideSearchResults() {
        this.searchResultsContainer.classList.add('hidden');
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text || '';
        return div.innerHTML;
    }
}

// ==========================================
// APPLICATION CONTROLLER
// Orchestrates all services
// ==========================================
class ScheduleApp {
    constructor() {
        // Services
        this.actionToastService = new ActionToastService();
        this.fileManager = new FileManager(
            document.getElementById('file-upload'),
            document.getElementById('remove-file-btn'),
            document.getElementById('file-name')
        );
        this.scheduleManager = new ScheduleManager();
        this.searchService = new SearchService([]);
        this.uiRenderer = new UIRenderer();

        // UI Elements
        this.searchBtn = document.getElementById('search-course-btn');
        this.clearAllBtn = document.getElementById('clear-all-btn');
        this.courseInput = document.getElementById('course-input');

        this.initialize();
    }

    initialize() {
        this.fileManager.initialize((data) => this.onFileLoaded(data));
        this.setupEventListeners();
    }

    setupEventListeners() {
        this.searchBtn.addEventListener('click', () => this.performSearch());
        this.clearAllBtn.addEventListener('click', () => this.clearSchedule());
        this.courseInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') this.performSearch();
        });
    }

    onFileLoaded(data) {
        this.searchService = new SearchService(data);
        this.uiRenderer.showCourseSection();
        this.actionToastService.show('✓ File loaded successfully!');
    }

    performSearch() {
        const searchQuery = this.courseInput.value.trim();
        if (!searchQuery) {
            this.actionToastService.show('⚠ Please enter a course code');
            return;
        }

        const results = this.searchService.search(searchQuery);

        if (results.length === 0) {
            this.actionToastService.show(`✗ Course "${searchQuery}" not found`);
            this.uiRenderer.hideSearchResults();
            return;
        }

        const rendered = this.uiRenderer.renderSearchResults(results, (index) => this.addCourseToSchedule(results, index));
        if (rendered) {
            this.actionToastService.show(`Found ${results.length} section(s)`);
        }
    }

    addCourseToSchedule(results, index) {
        const selectedCourse = results[index];
        const result = this.scheduleManager.addCourse(selectedCourse);

        if (result.success) {
            this.updateScheduleDisplay();
            this.uiRenderer.clearSearchInput();
            this.uiRenderer.hideSearchResults();
            this.actionToastService.show(`✓ ${selectedCourse.courseCode} added to schedule!`);
        } else {
            this.actionToastService.show(result.message);
        }
    }

    updateScheduleDisplay() {
        const sortedSchedule = this.scheduleManager.getScheduleSorted();
        this.uiRenderer.renderSchedule(sortedSchedule, (index) => this.removeCourseFromSchedule(index));
    }

    removeCourseFromSchedule(index) {
        const schedule = this.scheduleManager.getSchedule();
        const courseCode = schedule[index]?.courseCode;
        this.scheduleManager.removeCourse(index);
        this.updateScheduleDisplay();
        if (courseCode) {
            this.actionToastService.show(`✓ ${courseCode} removed from schedule`);
        }
    }

    clearSchedule() {
        if (this.scheduleManager.getSchedule().length === 0) {
            this.actionToastService.show('Schedule is already empty');
            return;
        }
        this.scheduleManager.clearSchedule();
        this.updateScheduleDisplay();
        this.actionToastService.show('✓ Schedule cleared');
    }
}

// ==========================================
// APPLICATION INITIALIZATION
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
    new ScheduleApp();
});