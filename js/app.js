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

        // ========== 座席表設定 ==========
        const seatColsInput = document.getElementById('seatCols');
        const seatRowsInput = document.getElementById('seatRows');
        const seatOrderRadios = document.querySelectorAll('input[name="seatOrder"]');

        if (seatColsInput) seatColsInput.addEventListener('input', () => this._updateSeatInfo());
        if (seatRowsInput) seatRowsInput.addEventListener('input', () => this._updateSeatInfo());
        
        seatOrderRadios.forEach(radio => {
            radio.addEventListener('change', () => this._updateSeatInfo());
        });

        const generateSeatingBtn = document.getElementById('generateSeatingBtn');
        if (generateSeatingBtn) {
            generateSeatingBtn.addEventListener('click', () => this._generateSeatingChart());
        }

        const previewSeatingBtn = document.getElementById('previewSeatingBtn');
        if (previewSeatingBtn) {
            previewSeatingBtn.addEventListener('click', () => this._previewSeating());
        }

        // ========== 課題採点表設定 ==========
        const numAssignmentsInput = document.getElementById('numAssignments');
        if (numAssignmentsInput) {
            numAssignmentsInput.addEventListener('input', (e) => {
                this.config.numAssignments = parseInt(e.target.value) || 10;
            });
        }

        const generateGradingBtn = document.getElementById('generateGradingBtn');
        if (generateGradingBtn) {
            generateGradingBtn.addEventListener('click', () => this._generateGradingSheet());
        }

        const previewGradingBtn = document.getElementById('previewGradingBtn');
        if (previewGradingBtn) {
            previewGradingBtn.addEventListener('click', () => this._previewGrading());
        }

        // ========== 出席票設定 ==========
        const numWeeksInput = document.getElementById('numWeeks');
        if (numWeeksInput) {
            numWeeksInput.addEventListener('input', (e) => {
                this.config.numWeeks = parseInt(e.target.value) || 15;
            });
        }

        const startDateInput = document.getElementById('startDate');
        if (startDateInput) {
            startDateInput.addEventListener('change', (e) => {
                this.config.startDate = e.target.value;
            });
        }

        const generateAttendanceBtn = document.getElementById('generateAttendanceBtn');
        if (generateAttendanceBtn) {
            generateAttendanceBtn.addEventListener('click', () => this._generateAttendanceSheet());
        }

        const previewAttendanceBtn = document.getElementById('previewAttendanceBtn');
        if (previewAttendanceBtn) {
            previewAttendanceBtn.addEventListener('click', () => this._previewAttendance());
        }

        // ========== サイドバーメニュー ==========
        const menuItems = document.querySelectorAll('.menu-item');
        menuItems.forEach(item => {
            item.addEventListener('click', () => {
                const tool = item.getAttribute('data-tool');
                this._switchTool(tool);
                
                // アクティブ状態を更新
                menuItems.forEach(m => m.classList.remove('active'));
                item.classList.add('active');
            });
        });

        // リセットボタン
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

    _switchTool(tool) {
        if (!this.students.length) {
            this._showStatus('uploadStatus', 'warning', '⚠️ 先に名簿をアップロードしてください');
            return;
        }

        // すべてのセクションを非表示
        document.getElementById('uploadSection').style.display = 'none';
        document.getElementById('seatingContent').style.display = 'none';
        document.getElementById('gradingContent').style.display = 'none';
        document.getElementById('attendanceContent').style.display = 'none';
        document.getElementById('completeSection').style.display = 'none';

        // 選択されたツールを表示
        switch (tool) {
            case 'seating':
                document.getElementById('seatingContent').style.display = 'block';
                this._updateSeatInfo();
                break;
            case 'grading':
                document.getElementById('gradingContent').style.display = 'block';
                break;
            case 'attendance':
                document.getElementById('attendanceContent').style.display = 'block';
                break;
        }
    }

    _updateSeatInfo() {
        const cols = parseInt(document.getElementById('seatCols').value) || 10;
        const rows = parseInt(document.getElementById('seatRows').value) || 10;
        const totalSeats = cols * rows;
        const seatOrder = document.querySelector('input[name="seatOrder"]:checked').value;

        this.config.cols = cols;
        this.config.rows = rows;
        this.config.seatOrder = seatOrder;

        const orderText = seatOrder === 'student-id' ? '学籍番号順' : 'ランダム';
        const statusText = `${cols}列 × ${rows}行 = ${totalSeats}座席 | 配置順序：${orderText} | 学生数：${this.students.length}名`;
        
        const seatInfo = document.getElementById('seatInfo');
        if (seatInfo) {
            if (totalSeats >= this.students.length) {
                seatInfo.textContent = statusText;
                seatInfo.style.color = '#28a745';
            } else {
                seatInfo.textContent = `⚠️ ${statusText} (座席が不足しています: ${totalSeats - this.students.length}人)`;
                seatInfo.style.color = '#dc3545';
            }
        }
    }

    _generateSeatingChart() {
        if (!this.students.length) {
            this._showStatus('seatingStatus', 'error', '❌ 先に名簿をアップロードしてください');
            return;
        }

        try {
            const cols = parseInt(document.getElementById('seatCols').value) || 10;
            const rows = parseInt(document.getElementById('seatRows').value) || 10;
            const seatOrder = document.querySelector('input[name="seatOrder"]:checked').value;

            const config = {
                cols: cols,
                rows: rows,
                seatOrder: seatOrder
            };

            ExcelGenerator.generateSeatingChart(this.students, config);
            this._showStatus('seatingStatus', 'success', '✅ 座席表を生成しました');
        } catch (error) {
            this._showStatus('seatingStatus', 'error', `❌ エラー：${error.message}`);
        }
    }

    _previewSeating() {
        const cols = parseInt(document.getElementById('seatCols').value) || 10;
        const rows = parseInt(document.getElementById('seatRows').value) || 10;
        const totalSeats = cols * rows;

        if (totalSeats < this.students.length) {
            alert(`座席数（${totalSeats}）が学生数（${this.students.length}）より少なくなっています`);
        } else {
            alert(`座席表プレビュー:\n${cols}列 × ${rows}行 = ${totalSeats}座席\n学生数: ${this.students.length}名`);
        }
    }

    _generateGradingSheet() {
        if (!this.students.length) {
            this._showStatus('gradingStatus', 'error', '❌ 先に名簿をアップロードしてください');
            return;
        }

        try {
            const numAssignments = parseInt(document.getElementById('numAssignments').value) || 10;
            const config = { numAssignments: numAssignments };

            ExcelGenerator.generateGradingSheet(this.students, config);
            this._showStatus('gradingStatus', 'success', '✅ 課題採点シートを生成しました');
        } catch (error) {
            this._showStatus('gradingStatus', 'error', `❌ エラー：${error.message}`);
        }
    }

    _previewGrading() {
        const numAssignments = parseInt(document.getElementById('numAssignments').value) || 10;
        alert(`採点シートプレビュー:\n課題数: ${numAssignments}\n学生数: ${this.students.length}名\n\n※シートの列は課題数に応じて動的に変更されます`);
    }

    _generateAttendanceSheet() {
        if (!this.students.length) {
            this._showStatus('attendanceStatus', 'error', '❌ 先に名簿をアップロードしてください');
            return;
        }

        try {
            const numWeeks = parseInt(document.getElementById('numWeeks').value) || 15;
            const startDate = document.getElementById('startDate').value;

            if (!startDate) {
                this._showStatus('attendanceStatus', 'error', '❌ 開始日を指定してください');
                return;
            }

            const config = { numWeeks: numWeeks, startDate: startDate };
            ExcelGenerator.generateAttendanceSheet(this.students, config);
            this._showStatus('attendanceStatus', 'success', '✅ 出席票を生成しました');
        } catch (error) {
            this._showStatus('attendanceStatus', 'error', `❌ エラー：${error.message}`);
        }
    }

    _previewAttendance() {
        const numWeeks = parseInt(document.getElementById('numWeeks').value) || 15;
        const startDate = document.getElementById('startDate').value;
        alert(`出席票プレビュー:\n週数: ${numWeeks}週\n開始日: ${startDate}\n学生数: ${this.students.length}名`);
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
        // 未処理のファイル数をカウント
        let unprocessedCount = 0;
        this.uploadedFiles.forEach(fileData => {
            if (fileData.studentCount === 0) {
                unprocessedCount++;
            }
        });

        if (unprocessedCount === 0) {
            // 全ファイルが既に処理済み
            this._updateFilesList();
            this._updateStudentCount();
            return;
        }

        let filesProcessed = 0;

        this.uploadedFiles.forEach((fileData, idx) => {
            // 既に処理済みのファイルはスキップ
            if (fileData.studentCount > 0) {
                return;
            }

            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const csv = event.target.result;
                    const csvData = parseCSV(csv);
                    const students = extractStudents(csvData);

                    fileData.students = students;
                    fileData.studentCount = students.length;

                    filesProcessed++;

                    // 全ファイルが処理完了したら
                    if (filesProcessed === unprocessedCount) {
                        this._consolidateStudents();
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

    _consolidateStudents() {
        // 全ファイルから学生データを統合
        let allStudents = [];
        
        this.uploadedFiles.forEach(fileData => {
            if (fileData.students && fileData.students.length > 0) {
                allStudents = [...allStudents, ...fileData.students];
            }
        });

        // 重複排除
        const uniqueStudents = [];
        const seen = new Set();
        
        for (const student of allStudents) {
            if (!seen.has(student.id)) {
                uniqueStudents.push(student);
                seen.add(student.id);
            }
        }

        // ソート
        uniqueStudents.sort((a, b) => a.id.localeCompare(b.id));
        
        this.students = uniqueStudents;
        this._updateFilesList();
        this._updateStudentCount();
        this._showStatus('uploadStatus', 'success', `✅ 合計${this.students.length}名の学生を読み込みました`);
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
