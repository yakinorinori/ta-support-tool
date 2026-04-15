/**
 * ユーティリティ関数
 */

/**
 * CSVファイルを解析してデータを取得
 */
function parseCSV(text) {
    const lines = text.trim().split('\n');
    const data = [];
    
    lines.forEach(line => {
        const row = [];
        let current = '';
        let insideQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                insideQuotes = !insideQuotes;
            } else if (char === ',' && !insideQuotes) {
                row.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        row.push(current.trim());
        
        if (row.some(cell => cell.length > 0)) {
            data.push(row);
        }
    });
    
    return data;
}

/**
 * 名簿CSVから学籍番号と名前を抽出
 */
function extractStudents(csvData) {
    // スキップ行がある場合を想定（先頭6行をスキップ）
    let startIdx = 0;
    for (let i = 0; i < csvData.length; i++) {
        if (csvData[i].some(cell => cell.includes('学籍番号'))) {
            startIdx = i + 1;
            break;
        }
    }
    
    // ヘッダー行がない場合は先頭から
    if (startIdx === 0) startIdx = 0;
    
    const students = [];
    
    for (let i = startIdx; i < csvData.length; i++) {
        const row = csvData[i];
        if (row.length >= 2) {
            const studentId = row[0]?.trim();
            const studentName = row[1]?.trim();
            
            if (studentId && studentName && studentId !== '学籍番号' && studentName !== '学生氏名') {
                students.push({
                    id: String(studentId),
                    name: studentName
                });
            }
        }
    }
    
    // 学籍番号でソート
    students.sort((a, b) => a.id.localeCompare(b.id));
    
    // 重複排除
    const unique = [];
    const seen = new Set();
    for (const student of students) {
        if (!seen.has(student.id)) {
            unique.push(student);
            seen.add(student.id);
        }
    }
    
    return unique;
}

/**
 * シャッフル（Fisher-Yates）
 */
function shuffle(array) {
    const arr = [...array];
    for (let i = arr.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [arr[i], arr[j]] = [arr[j], arr[i]];
    }
    return arr;
}

/**
 * 日付をYYYY-MM-DD形式にフォーマット
 */
function formatDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

/**
 * 日付を人間が読める形式にフォーマット
 */
function formatDateJP(date) {
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    return `${month}/${day}`;
}

/**
 * セルの背景色を設定（Excel互換）
 */
function setCellBackground(cell, color) {
    cell.fill = {
        type: 'solid',
        fgColor: { rgb: color }
    };
}

/**
 * セルのテキスト書式を設定
 */
function setCellStyle(cell, options = {}) {
    if (options.bold) {
        cell.font = { bold: true, ...cell.font };
    }
    if (options.fontSize) {
        cell.font = { size: options.fontSize, ...cell.font };
    }
    if (options.alignment) {
        cell.alignment = options.alignment;
    }
    if (options.border) {
        const border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
        };
        cell.border = border;
    }
    if (options.color) {
        cell.font = { color: { rgb: options.color }, ...cell.font };
    }
}

/**
 * 結果をダウンロード
 */
function downloadFile(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

/**
 * Excelワークブックをバイナリ文字列に変換
 */
function writeXLSXToBase64(workbook) {
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    return new Blob([wbout], { type: 'application/octet-stream' });
}
