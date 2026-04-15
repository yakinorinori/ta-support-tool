/**
 * TA Support Tool - メインアプリケーション
 */

class TAToolApp {
    constructor() {
        this.currentStep = 1;
        this.students = [];
        this.uploadedFiles = []; // ファイル管理用
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
        this.generatedFiles = [];
        this.init();
    }

    init() {
        this._setupEventListeners();
        this._setDefaultDate();
    }

    _setupEventListeners() {
        // CSVアップロード（メイン）
        const csvFile = document.getElementById('csvFiles');
        if (csvFile) {
            csvFile.addEventListener('change', (e) => this._handleMultipleCSVUpload(e));
        }

        // CSVアップロード（サイドバー）
        const sidebarCsvFiles = document.getElementById('sidebarCsvFiles');
        if (sidebarCsvFiles) {
            sidebarCsvFiles.addEventListener('change', (e) => this._handleSidebarCSVUpload(e));
        }

        // ファイルクリアボタン（サイドバー）
        const clearFilesBtn = document.getElementById('sidebarClearFilesBtn');
        if (clearFilesBtn) {
            clearFilesBtn.addEventListener('click', () => this._clearAllFiles());
        }

        // ファイルクリアボタン（メイン）
        const mainClearFilesBtn = document.getElementById('clearFilesBtn');
        if (mainClearFilesBtn) {
            mainClearFilesBtn.addEventListener('click', () => this._clearAllFiles());
        }

        // レイアウト選択
        const layoutRadios = document.querySelectorAll('input[name="layout"]');
        layoutRadios.forEach(radio => {
            radio.addEventListener('change', (e) => this._handleLayoutChange(e));
        });

        // 座席数の動的計算
        const colsInput = document.getElementById('cols');
        const rowsInput = document.getElementById('rows');
        const threeInput = document.getElementById('threeSeatRows');
        const twoInput = document.getElementById('twoSeatRows');

        if (colsInput) colsInput.addEventListener('input', () => this._updateSeatInfo());
        if (rowsInput) rowsInput.addEventListener('input', () => this._updateSeatInfo());
        if (threeInput) threeInput.addEventListener('input', () => this._updateSeatInfo());
        if (twoInput) twoInput.addEventListener('input', () => this._updateSeatInfo());

        // 出力ファイル選択
        const generateGradingCheckbox = document.getElementById('generate-grading');
        const generateAttendanceCheckbox = document.getElementById('generate-attendance');
        
        if (generateGradingCheckbox) {
            generateGradingCheckbox.addEventListener('change', (e) => {
                const gradingConfig = document.getElementById('gradingConfig');
                if (gradingConfig) gradingConfig.style.display = e.target.checked ? 'block' : 'none';
            });
        }
        if (generateAttendanceCheckbox) {
            generateAttendanceCheckbox.addEventListener('change', (e) => {
                const attendanceConfig = document.getElementById('attendanceConfig');
                if (attendanceConfig) attendanceConfig.style.display = e.target.checked ? 'block' : 'none';
            });
        }

        // 生成ボタン
        const generateBtn = document.getElementById('generateBtn');
        if (generateBtn) {
            generateBtn.addEventListener('click', () => this._generateFiles());
        }

        // リセットボタン
        const resetBtn = document.getElementById('resetBtn');
        if (resetBtn) {
            resetBtn.addEventListener('click', () => this._reset());
        }

        // サイドバーのリセットボタン
        const resetAllBtn = document.getElementById('resetAllBtn');
        if (resetAllBtn) {
            resetAllBtn.addEventListener('click', () => this._reset());
        }
    }

    _setDefaultDate() {
        const today = new Date();
        const dateInput = document.getElementById('startDate');
        dateInput.value = formatDate(today);
        this.config.startDate = formatDate(today);
    }

    _handleMultipleCSVUpload(e) {
        const files = e.target.files;
        if (!files || files.length === 0) return;

        // ファイルを追加
        this._addFilesToUploadList(Array.from(files));
    }

    _handleSidebarCSVUpload(e) {
        const files = e.target.files;
        if (!files || files.length === 0) return;

        // ファイルを追加
        this._addFilesToUploadList(Array.from(files));
    }

    _addFilesToUploadList(files) {
        const newFiles = files.filter(file => {
            // 同じ名前のファイルが既に存在するかチェック
            return !this.uploadedFiles.find(f => f.name === file.name);
        });

        // 新しいファイルを追加
        newFiles.forEach(file => {
            this.uploadedFiles.push({
                name: file.name,
                file: file,
                studentCount: 0,
                students: []
            });
        });

        // リスト表示を更新
        this._updateFilesList();

        // 全ファイルのデータを読み込む
        this._loadAllUploadedFiles();
    }

    _updateFilesList() {
        const sidebarList = document.getElementById('sidebarFilesListItems');
        const mainList = document.getElementById('filesList');

        if (!this.uploadedFiles.length) {
            // ファイルが無い場合は表示を隠す
            if (document.getElementById('uploadedFilesList')) {
                document.getElementById('uploadedFilesList').style.display = 'none';
            }
            if (document.getElementById('sidebarFilesList')) {
                document.getElementById('sidebarFilesList').style.display = 'none';
            }
            return;
        }

        // ファイル一覧表示を更新
        let html = '';
        this.uploadedFiles.forEach((fileData, idx) => {
            html += `<li data-index="${idx}">📄 ${fileData.name} <span class="file-count">(${fileData.studentCount}名)</span></li>`;
        });

        if (sidebarList) sidebarList.innerHTML = html;
        if (mainList) mainList.innerHTML = html;

        // コンテナを表示
        if (document.getElementById('uploadedFilesList')) {
            document.getElementById('uploadedFilesList').style.display = 'block';
        }
        if (document.getElementById('sidebarFilesList')) {
            document.getElementById('sidebarFilesList').style.display = 'block';
        }
    }

    _loadAllUploadedFiles() {
        let totalStudents = [];
        let filesProcessed = 0;

        this.uploadedFiles.forEach((fileData, idx) => {
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const csv = event.target.result;
                    const csvData = parseCSV(csv);
                    const students = extractStudents(csvData);

                    fileData.students = students;
                    fileData.studentCount = students.length;
                    totalStudents = [...totalStudents, ...students];

                    filesProcessed++;

                    // 全ファイルが処理完了したら
                    if (filesProcessed === this.uploadedFiles.length) {
                        this.students = totalStudents;
                        this._updateFilesList();
                        this._updateStudentCount();
                        this._showStatus('uploadStatus', 'success', `✅ 合計${this.students.length}名の学生を読み込みました`);
                    }
                } catch (error) {
                    this._showStatus('uploadStatus', 'error', `❌ ${fileData.name}: ${error.message}`);
                }
            };

            reader.onerror = () => {
                this._showStatus('uploadStatus', 'error', `❌ ${fileData.name}: ファイルが読み込めません`);
            };

            reader.readAsText(fileData.file);
        });
    }

    _updateStudentCount() {
        const studentCountElement = document.querySelector('.student-count');
        if (studentCountElement) {
            studentCountElement.textContent = `${this.students.length}名`;
        }
    }

    _clearAllFiles() {
        // ファイル入力をクリア
        const csvFiles = document.getElementById('csvFiles');
        const sidebarCsvFiles = document.getElementById('sidebarCsvFiles');
        if (csvFiles) csvFiles.value = '';
        if (sidebarCsvFiles) sidebarCsvFiles.value = '';

        // データをクリア
        this.uploadedFiles = [];
        this.students = [];
        this._updateFilesList();
        this._updateStudentCount();

        // ステータスメッセージをクリア
        const uploadStatus = document.getElementById('uploadStatus');
        if (uploadStatus) uploadStatus.textContent = '';

        // 表示を隠す
        const uploadedFilesList = document.getElementById('uploadedFilesList');
        const sidebarFilesList = document.getElementById('sidebarFilesList');
        if (uploadedFilesList) uploadedFilesList.style.display = 'none';
        if (sidebarFilesList) sidebarFilesList.style.display = 'none';
    }

    _handleCSVUpload(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const csv = event.target.result;
                const csvData = parseCSV(csv);
                this.students = extractStudents(csvData);

                if (this.students.length === 0) {
                    this._showStatus('uploadStatus', 'error', '❌ 名簿を読み込めませんでした。CSVの形式を確認してください。');
                    return;
                }

                this._showStatus('uploadStatus', 'success', `✅ ${this.students.length}名の学生を読み込みました`);
                this._goToStep(2);
            } catch (error) {
                this._showStatus('uploadStatus', 'error', `❌ エラー：${error.message}`);
            }
        };
        reader.readAsText(file);
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

    _generateFiles() {
        try {
            const generateSeating = document.getElementById('generate-seating').checked;
            const generateGrading = document.getElementById('generate-grading').checked;
            const generateAttendance = document.getElementById('generate-attendance').checked;

            if (!generateSeating && !generateGrading && !generateAttendance) {
                this._showStatus('generationStatus', 'warning', '⚠️ 最低1つのファイルを選択してください');
                return;
            }

            this.generatedFiles = [];
            const status = document.getElementById('generationStatus');
            status.innerHTML = '<span class="spinner"></span>ファイルを生成中...';
            status.className = 'status-message';

            // 非同期で実行
            setTimeout(() => {
                try {
                    if (generateSeating) {
                        ExcelGenerator.generateSeatingChart(this.students, this.config.layout, this.config);
                        this.generatedFiles.push('座席表.xlsx');
                    }

                    if (generateGrading) {
                        ExcelGenerator.generateGradingSheet(this.students, this.config);
                        this.generatedFiles.push('課題採点シート.xlsx');
                    }

                    if (generateAttendance) {
                        ExcelGenerator.generateAttendanceSheet(this.students, this.config);
                        this.generatedFiles.push('出席票.xlsx');
                    }

                    status.innerHTML = '';
                    status.className = 'status-message success';
                    this._goToStep(5);
                } catch (error) {
                    status.textContent = `❌ エラーが発生しました：${error.message}`;
                    status.className = 'status-message error';
                    console.error(error);
                }
            }, 100);
        } catch (error) {
            this._showStatus('generationStatus', 'error', `❌ エラー：${error.message}`);
            console.error(error);
        }
    }

    _goToStep(step) {
        // 前のステップを非表示
        for (let i = 1; i <= 5; i++) {
            const stepElement = document.getElementById(`step${i}`);
            if (stepElement) {
                stepElement.style.display = 'none';
            }
        }

        // 新しいステップを表示
        const newStep = document.getElementById(`step${step}`);
        if (newStep) {
            newStep.style.display = 'block';
            newStep.scrollIntoView({ behavior: 'smooth' });
        }

        this.currentStep = step;

        // 次のステップへの遷移を有効化
        if (step === 2) {
            this._updateSeatInfo();
        } else if (step === 3) {
            document.getElementById('step3').style.display = 'block';
            this._goToStep(4); // ステップ3は自動スキップで、ステップ4へ
            return;
        }
    }

    _showStatus(elementId, type, message) {
        const element = document.getElementById(elementId);
        element.textContent = message;
        element.className = `status-message ${type}`;
    }

    _displayGeneratedFiles() {
        const list = document.getElementById('downloadList');
        list.innerHTML = '';

        this.generatedFiles.forEach(filename => {
            const li = document.createElement('li');
            li.innerHTML = `
                <span>📄 ${filename}</span>
                <span style="color: #28a745; font-weight: bold;">✅ ダウンロード完了</span>
            `;
            list.appendChild(li);
        });
    }

    _reset() {
        // UIをリセット
        const csvFiles = document.getElementById('csvFiles');
        const sidebarCsvFiles = document.getElementById('sidebarCsvFiles');
        if (csvFiles) csvFiles.value = '';
        if (sidebarCsvFiles) sidebarCsvFiles.value = '';

        // ステータスメッセージをリセット
        const uploadStatus = document.getElementById('uploadStatus');
        const generateGradingCheckbox = document.getElementById('generate-grading');
        const generateAttendanceCheckbox = document.getElementById('generate-attendance');
        const gradingConfig = document.getElementById('gradingConfig');
        const attendanceConfig = document.getElementById('attendanceConfig');
        const generationStatus = document.getElementById('generationStatus');
        const downloadList = document.getElementById('downloadList');

        if (uploadStatus) uploadStatus.textContent = '';
        if (generateGradingCheckbox) generateGradingCheckbox.checked = false;
        if (generateAttendanceCheckbox) generateAttendanceCheckbox.checked = false;
        if (gradingConfig) gradingConfig.style.display = 'none';
        if (attendanceConfig) attendanceConfig.style.display = 'none';
        if (generationStatus) generationStatus.textContent = '';
        if (downloadList) downloadList.innerHTML = '';

        // データをリセット
        this.uploadedFiles = [];
        this.students = [];
        this.generatedFiles = [];

        // UI表示をリセット
        this._updateFilesList();
        this._updateStudentCount();

        // ステップ1に戻る
        this._goToStep(1);
    }
}

// ページ読み込み時にアプリケーションを初期化
document.addEventListener('DOMContentLoaded', () => {
    window.app = new TAToolApp();
});
