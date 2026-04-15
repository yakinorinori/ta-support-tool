/**
 * Excel ファイル生成ロジック
 */

class ExcelGenerator {
    /**
     * 座席表を生成
     */
    static generateSeatingChart(students, layout = 'grid', config = {}) {
        const workbook = XLSX.utils.book_new();
        
        if (layout === 'grid') {
            this._generateGridSeating(workbook, students, config);
        } else if (layout === 'mixed') {
            this._generateMixedSeating(workbook, students, config);
        }
        
        XLSX.writeFile(workbook, '座席表.xlsx');
        return true;
    }

    /**
     * グリッド座席配置を生成
     */
    static _generateGridSeating(workbook, students, config) {
        const ws = XLSX.utils.aoa_to_sheet([]);
        const cols = parseInt(config.cols) || 10;
        const rows = parseInt(config.rows) || 15;
        
        // タイトル
        ws['A1'] = '座席表';
        ws['A1'].s = { font: { bold: true, sz: 14 } };
        
        // ヘッダー行
        for (let c = 0; c < cols; c++) {
            const cellRef = XLSX.utils.encode_cell({ r: 1, c });
            ws[cellRef] = `列${c + 1}`;
            ws[cellRef].s = {
                fill: { fgColor: { rgb: '4472C4' } },
                font: { bold: true, color: { rgb: 'FFFFFF' } },
                alignment: { horizontal: 'center', vertical: 'center' }
            };
        }
        
        // 学籍番号を配置
        const shuffled = shuffle(students);
        let idx = 0;
        for (let col = 0; col < cols; col++) {
            for (let row = 0; row < rows; row++) {
                if (idx < shuffled.length) {
                    const cellRef = XLSX.utils.encode_cell({ r: row + 2, c: col });
                    ws[cellRef] = shuffled[idx].id;
                    ws[cellRef].s = {
                        alignment: { horizontal: 'center', vertical: 'center' },
                        border: {
                            top: { style: 'thin' },
                            bottom: { style: 'thin' },
                            left: { style: 'thin' },
                            right: { style: 'thin' }
                        }
                    };
                    idx++;
                }
            }
        }
        
        // 列幅と行高さを設定
        ws['!cols'] = Array(cols).fill({ wch: 12 });
        ws['!rows'] = Array(rows + 2).fill({ hpx: 25 });
        
        workbook.SheetNames.push('座席表');
        workbook.Sheets['座席表'] = ws;
    }

    /**
     * 混合座席配置（3人席+2人席）を生成
     */
    static _generateMixedSeating(workbook, students, config) {
        const ws = XLSX.utils.aoa_to_sheet([]);
        const threeRows = parseInt(config.threeSeatRows) || 5;
        const twoRows = parseInt(config.twoSeatRows) || 10;
        
        // タイトル
        ws['A1'] = '座席表（3人席+2人席）';
        ws['A1'].s = { font: { bold: true, sz: 14 } };
        
        // シャッフル
        const shuffled = shuffle(students);
        
        let currentRow = 2;
        let idx = 0;
        
        // 3人席セクション
        ws[`A${currentRow}`] = '【3人席】5列';
        ws[`A${currentRow}`].s = { font: { bold: true, sz: 11 } };
        currentRow++;
        
        for (let r = 0; r < threeRows; r++) {
            for (let c = 0; c < 3; c++) {
                if (idx < shuffled.length) {
                    const cellRef = XLSX.utils.encode_cell({ r: currentRow - 1 + r, c });
                    ws[cellRef] = shuffled[idx].id;
                    ws[cellRef].s = {
                        alignment: { horizontal: 'center', vertical: 'center' },
                        border: {
                            top: { style: 'thin' },
                            bottom: { style: 'thin' },
                            left: { style: 'thin' },
                            right: { style: 'thin' }
                        }
                    };
                    idx++;
                }
            }
        }
        
        currentRow += threeRows + 1;
        
        // 2人席セクション
        ws[`A${currentRow}`] = '【2人席】10列';
        ws[`A${currentRow}`].s = { font: { bold: true, sz: 11 } };
        currentRow++;
        
        for (let r = 0; r < twoRows; r++) {
            for (let c = 0; c < 2; c++) {
                if (idx < shuffled.length) {
                    const cellRef = XLSX.utils.encode_cell({ r: currentRow - 1 + r, c });
                    ws[cellRef] = shuffled[idx].id;
                    ws[cellRef].s = {
                        alignment: { horizontal: 'center', vertical: 'center' },
                        border: {
                            top: { style: 'thin' },
                            bottom: { style: 'thin' },
                            left: { style: 'thin' },
                            right: { style: 'thin' }
                        }
                    };
                    idx++;
                }
            }
        }
        
        ws['!cols'] = [{ wch: 12 }, { wch: 12 }, { wch: 12 }];
        
        workbook.SheetNames.push('座席表');
        workbook.Sheets['座席表'] = ws;
    }

    /**
     * 課題採点シートを生成
     */
    static generateGradingSheet(students, config = {}) {
        const workbook = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([]);
        
        const numAssignments = parseInt(config.numAssignments) || 10;
        
        // タイトル
        ws['A1'] = '課題採点シート';
        ws['A1'].s = { font: { bold: true, sz: 14 } };
        
        // ヘッダー
        const headers = ['学籍番号', '氏名'];
        for (let i = 0; i < numAssignments; i++) {
            headers.push(`課題${i + 1}`);
        }
        headers.push('合計');
        
        for (let c = 0; c < headers.length; c++) {
            const cellRef = XLSX.utils.encode_cell({ r: 1, c });
            ws[cellRef] = headers[c];
            ws[cellRef].s = {
                fill: { fgColor: { rgb: '4472C4' } },
                font: { bold: true, color: { rgb: 'FFFFFF' } },
                alignment: { horizontal: 'center', vertical: 'center' }
            };
        }
        
        // データ行
        for (let r = 0; r < students.length; r++) {
            const student = students[r];
            const excelRow = r + 2;
            
            // 学籍番号
            const idCell = XLSX.utils.encode_cell({ r: excelRow - 1, c: 0 });
            ws[idCell] = student.id;
            ws[idCell].s = { alignment: { horizontal: 'center', vertical: 'center' } };
            
            // 氏名
            const nameCell = XLSX.utils.encode_cell({ r: excelRow - 1, c: 1 });
            ws[nameCell] = student.name;
            ws[nameCell].s = { alignment: { horizontal: 'left', vertical: 'center' } };
            
            // 課題欄
            for (let c = 0; c < numAssignments; c++) {
                const cellRef = XLSX.utils.encode_cell({ r: excelRow - 1, c: c + 2 });
                ws[cellRef] = '';
                ws[cellRef].s = { 
                    alignment: { horizontal: 'center', vertical: 'center' },
                    numFmt: '0'
                };
            }
            
            // 合計欄
            const totalCell = XLSX.utils.encode_cell({ r: excelRow - 1, c: numAssignments + 2 });
            const colStart = XLSX.utils.encode_col(2);
            const colEnd = XLSX.utils.encode_col(numAssignments + 1);
            ws[totalCell] = `=SUM(${colStart}${excelRow}:${colEnd}${excelRow})`;
            ws[totalCell].s = {
                fill: { fgColor: { rgb: 'FFF2CC' } },
                font: { bold: true },
                alignment: { horizontal: 'center', vertical: 'center' }
            };
        }
        
        // 列幅
        ws['!cols'] = [
            { wch: 12 },
            { wch: 18 },
            ...Array(numAssignments + 1).fill({ wch: 10 })
        ];
        
        workbook.SheetNames.push('課題採点');
        workbook.Sheets['課題採点'] = ws;
        
        XLSX.writeFile(workbook, '課題採点シート.xlsx');
        return true;
    }

    /**
     * 出席票を生成
     */
    static generateAttendanceSheet(students, config = {}) {
        const workbook = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([]);
        
        const numWeeks = parseInt(config.numWeeks) || 15;
        let startDate = new Date(config.startDate || '2026-04-21');
        if (isNaN(startDate)) startDate = new Date('2026-04-21');
        
        // タイトル
        ws['A1'] = '出席票';
        ws['A1'].s = { font: { bold: true, sz: 12 } };
        
        // ヘッダー
        ws['A2'] = '学籍番号';
        ws['A2'].s = {
            fill: { fgColor: { rgb: 'CCCCCC' } },
            font: { bold: true },
            alignment: { horizontal: 'center', vertical: 'center' }
        };
        
        ws['B2'] = '氏名';
        ws['B2'].s = {
            fill: { fgColor: { rgb: 'CCCCCC' } },
            font: { bold: true },
            alignment: { horizontal: 'center', vertical: 'center' }
        };
        
        // 日付ヘッダー
        for (let i = 0; i < numWeeks; i++) {
            const date = new Date(startDate);
            date.setDate(date.getDate() + i * 7);
            const cellRef = XLSX.utils.encode_cell({ r: 1, c: i + 2 });
            ws[cellRef] = `第${i + 1}回\n${formatDateJP(date)}`;
            ws[cellRef].s = {
                fill: { fgColor: { rgb: 'E7E6E6' } },
                font: { bold: true, sz: 9 },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
            };
        }
        
        // データ行
        for (let r = 0; r < students.length; r++) {
            const student = students[r];
            const excelRow = r + 2;
            
            // 学籍番号
            const idCell = XLSX.utils.encode_cell({ r: excelRow - 1, c: 0 });
            ws[idCell] = student.id;
            ws[idCell].s = {
                border: {
                    top: { style: 'thin' },
                    bottom: { style: 'thin' },
                    left: { style: 'thin' },
                    right: { style: 'thin' }
                },
                alignment: { horizontal: 'center', vertical: 'center' }
            };
            
            // 氏名
            const nameCell = XLSX.utils.encode_cell({ r: excelRow - 1, c: 1 });
            ws[nameCell] = student.name;
            ws[nameCell].s = {
                border: {
                    top: { style: 'thin' },
                    bottom: { style: 'thin' },
                    left: { style: 'thin' },
                    right: { style: 'thin' }
                },
                alignment: { horizontal: 'left', vertical: 'center' }
            };
            
            // チェック欄
            for (let c = 0; c < numWeeks; c++) {
                const cellRef = XLSX.utils.encode_cell({ r: excelRow - 1, c: c + 2 });
                ws[cellRef] = '';
                ws[cellRef].s = {
                    border: {
                        top: { style: 'thin' },
                        bottom: { style: 'thin' },
                        left: { style: 'thin' },
                        right: { style: 'thin' }
                    },
                    alignment: { horizontal: 'center', vertical: 'center' },
                    fill: { fgColor: { rgb: 'FFFFFF' } }
                };
            }
        }
        
        // 列幅
        ws['!cols'] = [
            { wch: 12 },
            { wch: 20 },
            ...Array(numWeeks).fill({ wch: 5 })
        ];
        ws['!rows'] = Array(students.length + 2).fill({ hpx: 25 });
        ws['!rows'][1].hpx = 30;
        
        workbook.SheetNames.push('出席票');
        workbook.Sheets['出席票'] = ws;
        
        XLSX.writeFile(workbook, '出席票.xlsx');
        return true;
    }
}
