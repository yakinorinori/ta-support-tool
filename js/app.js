/**
 * TA Support Tool - メインアプリケーション
 */

class TAToolApp {
    constructor() {
        this.currentStep = 1;
        this.students = [];
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
        // CSVアップロード
        const csvFile = document.getElementById('csvFile');
        csvFile.addEventListener('change', (e) => this._handleCSVUpload(e));

        // レイアウト選択
        const layoutRadios = document.querySelectorAll('input[name="layout"]');
        layoutRadios.forEach(radio => {
            radio.addEventListener('change', (e) => this._handleLayoutChange(e));
        });

        // 座席数の動的計算
        document.getElementById('cols').addEventListener('input', () => this._updateSeatInfo());
        document.getElementById('rows').addEventListener('input', () => this._updateSeatInfo());
        document.getElementById('threeSeatRows').addEventListener('input', () => this._updateSeatInfo());
        document.getElementById('twoSeatRows').addEventListener('input', () => this._updateSeatInfo());

        // 出力ファイル選択
        document.getElementById('generate-grading').addEventListener('change', (e) => {
            document.getElementById('gradingConfig').style.display = e.target.checked ? 'block' : 'none';
        });
        document.getElementById('generate-attendance').addEventListener('change', (e) => {
            document.getElementById('attendanceConfig').style.display = e.target.checked ? 'block' : 'none';
        });

        // 生成ボタン
        document.getElementById('generateBtn').addEventListener('click', () => this._generateFiles());

        // リセットボタン
        document.getElementById('resetBtn').addEventListener('click', () => this._reset());
    }

    _setDefaultDate() {
        const today = new Date();
        const dateInput = document.getElementById('startDate');
        dateInput.value = formatDate(today);
        this.config.startDate = formatDate(today);
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
        document.getElementById('csvFile').value = '';
        document.getElementById('uploadStatus').textContent = '';
        document.getElementById('generate-seating').checked = true;
        document.getElementById('generate-grading').checked = false;
        document.getElementById('generate-attendance').checked = false;
        document.getElementById('gradingConfig').style.display = 'none';
        document.getElementById('attendanceConfig').style.display = 'none';
        document.getElementById('generationStatus').textContent = '';
        document.getElementById('downloadList').innerHTML = '';

        // データをリセット
        this.students = [];
        this.generatedFiles = [];

        // ステップ1に戻る
        this._goToStep(1);
    }
}

// ページ読み込み時にアプリケーションを初期化
document.addEventListener('DOMContentLoaded', () => {
    window.app = new TAToolApp();
});
