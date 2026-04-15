/**
 * TA Support Tool v2.0 - 複数CSV対応、サイドバーナビゲーション
 */

class TAToolAppV2 {
    constructor() {
        this.students = [];
        this.uploadedFiles = [];
        this.currentTool = 'seating';
        this.config = {
            layout: 'grid',
            cols: 10,
            rows: 15,
            threeSeatRows: 5,
            twoSeatRows: 10,
            numAssignments: 10,
            numWeeks: 15,
            startDate: formatDate(new Date())
        };
        this.init();
    }

    init() {
        this._setupEventListeners();
        this._setDefaultDate();
    }

    _setupEventListeners() {
        // ファイルアップロード
        const fileUploadArea = document.getElementById('fileUploadArea');
        const csvFiles = document.getElementById('csvFiles');

        fileUploadArea.addEventListener('click', () => csvFiles.click());
        fileUploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            fileUploadArea.style.borderColor = 'var(--primary-color)';
        });
        fileUploadArea.addEventListener('dragleave', () => {
            fileUploadArea.style.borderColor = 'var(--border-color)';
        });
        fileUploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            fileUploadArea.style.borderColor = 'var(--border-color)';
            this._handleFileUpload(e.dataTransfer.files);
        });

        csvFiles.addEventListener('change', (e) => this._handleFileUpload(e.target.files));

        // サイドバーメニュー
        document.querySelectorAll('.menu-item').forEach(btn => {
            btn.addEventListener('click', (e) => this._switchTool(e.target.closest('.menu-item').dataset.tool));
        });

        // ファイルクリア
        document.getElementById('clearFilesBtn').addEventListener('click', () => this._clearFiles());

        // リセット
        document.getElementById('resetAllBtn').addEventListener('click', () => this._resetAll());

        // レイアウト選択
        document.querySelectorAll('input[name="layout"]').forEach(radio => {
            radio.addEventListener('change', (e) => this._handleLayoutChange(e));
        });

        // 座席数の動的計算
        document.getElementById('cols').addEventListener('input', () => this._updateSeatInfo());
        document.getElementById('rows').addEventListener('input', () => this._updateSeatInfo());
        document.getElementById('threeSeatRows').addEventListener('input', () => this._updateSeatInfo());
        document.getElementById('twoSeatRows').addEventListener('input', () => this._updateSeatInfo());

        // 生成ボタン
        document.getElementById('generateSeatingBtn').addEventListener('click', () => this._generateSeating());
        document.getElementById('generateGradingBtn').addEventListener('click', () => this._generateGrading());
        document.getElementById('generateAttendanceBtn').addEventListener('click', () => this._generateAttendance());
    }

    _setDefaultDate() {
        const today = new Date();
        const dateInput = document.getElementById('startDate');
        dateInput.value = formatDate(today);
        this.config.startDate = formatDate(today);
    }

    _handleFileUpload(files) {
        let validFiles = 0;
        let allStudents = [];
        const fileList = [];

        Array.from(files).forEach((file, index) => {
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const csv = event.target.result;
                    const csvData = parseCSV(csv);
                    const students = extractStudents(csvData);

                    if (students.length > 0) {
                        allStudents = allStudents.concat(students);
                        fileList.push({
                            name: file.name,
                            count: students.length
                        });
                        validFiles++;
                    }

                    // すべてのファイルが処理されたか確認
                    if (validFiles + (Object.keys(files).length - index - 1) < Object.keys(files).length) {
                        return;
                    }

                    if (validFiles === 0) {
                        this._showStatus('uploadStatus', 'error', '❌ 有効なCSVファイルが見つかりません');
                        return;
                    }

                    // 重複排除とソート
                    const uniqueStudents = [];
                    const seen = new Set();
                    for (const student of allStudents) {
                        if (!seen.has(student.id)) {
                            uniqueStudents.push(student);
                            seen.add(student.id);
                        }
                    }
                    uniqueStudents.sort((a, b) => a.id.localeCompare(b.id));

                    this.students = uniqueStudents;
                    this.uploadedFiles = fileList;

                    // UI更新
                    this._updateUploadedFilesList();
                    this._showStatus('uploadStatus', 'success', `✅ ${this.students.length}名の学生を読み込みました`);
                    this._updateStudentCount();
                } catch (error) {
                    this._showStatus('uploadStatus', 'error', `❌ エラー：${error.message}`);
                }
            };
            reader.readAsText(file);
        });
    }

    _updateUploadedFilesList() {
        const listContainer = document.getElementById('uploadedFilesList');
        const filesList = document.getElementById('filesList');

        filesList.innerHTML = '';
        this.uploadedFiles.forEach(file => {
            const li = document.createElement('li');
            li.innerHTML = `
                <span>📄 ${file.name}</span>
                <span class="badge">${file.count}名</span>
            `;
            filesList.appendChild(li);
        });

        listContainer.style.display = this.uploadedFiles.length > 0 ? 'block' : 'none';
    }

    _updateStudentCount() {
        document.querySelector('.student-count').textContent = `${this.students.length}名`;
    }

    _clearFiles() {
        this.students = [];
        this.uploadedFiles = [];
        document.getElementById('csvFiles').value = '';
        document.getElementById('uploadedFilesList').style.display = 'none';
        document.getElementById('uploadStatus').textContent = '';
        this._updateStudentCount();
    }

    _switchTool(tool) {
        if (this.students.length === 0) {
            this._showStatus('uploadStatus', 'warning', '⚠️ 先に名簿をアップロードしてください');
            return;
        }

        this.currentTool = tool;

        // メニューのアクティブ状態を更新
        document.querySelectorAll('.menu-item').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tool === tool);
        });

        // コンテンツセクションを切り替え
        document.querySelectorAll('.content-section').forEach(section => {
            section.classList.remove('active');
        });

        switch (tool) {
            case 'seating':
                document.getElementById('seatingContent').classList.add('active');
                this._updateSeatInfo();
                break;
            case 'grading':
                document.getElementById('gradingContent').classList.add('active');
                break;
            case 'attendance':
                document.getElementById('attendanceContent').classList.add('active');
                break;
        }
    }

    _handleLayoutChange(e) {
        const layout = e.target.value;
        this.config.layout = layout;

        const gridConfig = document.getElementById('gridConfig');
        const mixedConfig = document.getElementById('mixedConfig');

        if (layout === 'grid') {
            gridConfig.style.display = 'block';
            mixedConfig.style.display = 'none';
        } else {
            gridConfig.style.display = 'none';
            mixedConfig.style.display = 'block';
        }

        this._updateSeatInfo();
    }

    _updateSeatInfo() {
        const layout = this.config.layout;
        let totalSeats = 0;
        let info = '';

        if (layout === 'grid') {
            const cols = parseInt(document.getElementById('cols').value) || 10;
            const rows = parseInt(document.getElementById('rows').value) || 15;
            totalSeats = cols * rows;
            this.config.cols = cols;
            this.config.rows = rows;
            info = `グリッド配置：${cols}列 × ${rows}行 = ${totalSeats}座席`;
        } else {
            const threeRows = parseInt(document.getElementById('threeSeatRows').value) || 5;
            const twoRows = parseInt(document.getElementById('twoSeatRows').value) || 10;
            totalSeats = threeRows * 3 + twoRows * 2;
            this.config.threeSeatRows = threeRows;
            this.config.twoSeatRows = twoRows;
            info = `混合配置：3人席${threeRows}列(${threeRows * 3}座席) + 2人席${twoRows}列(${twoRows * 2}座席) = ${totalSeats}座席`;
        }

        const statusText = `${info} | 学生数：${this.students.length}名`;
        document.getElementById('seatInfo').textContent = statusText;
    }

    _generateSeating() {
        try {
            const status = document.getElementById('seatingStatus');
            status.innerHTML = '<span class="spinner"></span>座席表を生成中...';
            status.className = 'status-message';

            setTimeout(() => {
                try {
                    ExcelGenerator.generateSeatingChart(this.students, this.config.layout, this.config);
                    status.innerHTML = '';
                    status.className = 'status-message success';
                    status.textContent = '✅ 座席表をダウンロードしました';
                } catch (error) {
                    status.textContent = `❌ エラー：${error.message}`;
                    status.className = 'status-message error';
                    console.error(error);
                }
            }, 100);
        } catch (error) {
            this._showStatus('seatingStatus', 'error', `❌ エラー：${error.message}`);
        }
    }

    _generateGrading() {
        try {
            this.config.numAssignments = parseInt(document.getElementById('numAssignments').value) || 10;
            const status = document.getElementById('gradingStatus');
            status.innerHTML = '<span class="spinner"></span>課題採点シートを生成中...';
            status.className = 'status-message';

            setTimeout(() => {
                try {
                    ExcelGenerator.generateGradingSheet(this.students, this.config);
                    status.innerHTML = '';
                    status.className = 'status-message success';
                    status.textContent = '✅ 課題採点シートをダウンロードしました';
                } catch (error) {
                    status.textContent = `❌ エラー：${error.message}`;
                    status.className = 'status-message error';
                    console.error(error);
                }
            }, 100);
        } catch (error) {
            this._showStatus('gradingStatus', 'error', `❌ エラー：${error.message}`);
        }
    }

    _generateAttendance() {
        try {
            this.config.numWeeks = parseInt(document.getElementById('numWeeks').value) || 15;
            this.config.startDate = document.getElementById('startDate').value || formatDate(new Date());
            
            const status = document.getElementById('attendanceStatus');
            status.innerHTML = '<span class="spinner"></span>出席票を生成中...';
            status.className = 'status-message';

            setTimeout(() => {
                try {
                    ExcelGenerator.generateAttendanceSheet(this.students, this.config);
                    status.innerHTML = '';
                    status.className = 'status-message success';
                    status.textContent = '✅ 出席票をダウンロードしました';
                } catch (error) {
                    status.textContent = `❌ エラー：${error.message}`;
                    status.className = 'status-message error';
                    console.error(error);
                }
            }, 100);
        } catch (error) {
            this._showStatus('attendanceStatus', 'error', `❌ エラー：${error.message}`);
        }
    }

    _showStatus(elementId, type, message) {
        const element = document.getElementById(elementId);
        element.textContent = message;
        element.className = `status-message ${type}`;
    }

    _resetAll() {
        // UIをリセット
        this._clearFiles();
        document.getElementById('csvFiles').value = '';
        document.getElementById('uploadStatus').textContent = '';
        
        // 全セクションを非表示
        document.querySelectorAll('.content-section').forEach(section => {
            section.classList.remove('active');
        });
        
        // uploadSectionのみ表示
        document.getElementById('uploadSection').classList.add('active');

        // メニューを座席表にリセット
        document.querySelectorAll('.menu-item').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tool === 'seating');
        });

        this.currentTool = 'seating';
    }
}

// ページ読み込み時にアプリケーションを初期化
document.addEventListener('DOMContentLoaded', () => {
    window.app = new TAToolAppV2();
});
