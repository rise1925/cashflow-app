const { GoogleSpreadsheet } = require('google-spreadsheet');
const fs = require('fs');
const path = require('path');

class GoogleSheetsService {
    constructor() {
        this.doc = null;
        this.initialized = false;
    }

    /**
     * Google Sheets 초기화
     */
    async initialize() {
        try {
            if (this.initialized) return;

            const credentialsPath = path.join(
                __dirname,
                '../../',
                process.env.GOOGLE_CREDENTIALS_PATH || './credentials.json'
            );

            if (!fs.existsSync(credentialsPath)) {
                throw new Error(`credentials.json을 찾을 수 없습니다: ${credentialsPath}`);
            }

            const credentials = require(credentialsPath);

            // google-spreadsheet 3.x 방식
            this.doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID);

            // 서비스 계정으로 인증
            await this.doc.useServiceAccountAuth({
                client_email: credentials.client_email,
                private_key: credentials.private_key
            });

            await this.doc.loadInfo();
            this.initialized = true;
            console.log('✅ Google Sheets 연동 성공');
        } catch (error) {
            console.error('❌ Google Sheets 초기화 실패:', error.message);
            throw error;
        }
    }

    /**
     * 월별 데이터 조회
     */
    async getMonthlyData() {
        try {
            await this.initialize();

            // 첫 번째 시트
            const sheet = this.doc.sheetsByIndex[0];
            console.log('📄 시트 이름:', sheet.title);

            // 1. 첫 번째 행의 셀을 직접 로드하여 월 코드 찾기
            await sheet.loadCells('A1:AZ1');

            const monthColumns = {}; // { '01': 4, '02': 7, '03': 10, ... }
            for (let i = 0; i < sheet.columnCount; i++) {
                const cell = sheet.getCell(0, i);
                if (cell.value) {
                    const match = String(cell.value).match(/^(\d{2})/);
                    if (match) {
                        const monthCode = match[1];
                        monthColumns[monthCode] = i;
                        console.log(`📅 ${monthCode}월 발견 -> 열 인덱스: ${i}`);
                    }
                }
            }

            if (Object.keys(monthColumns).length === 0) {
                throw new Error('월 코드를 찾을 수 없습니다.');
            }

            const rows = await sheet.getRows();

            // 2. total 행에서 지출 찾기
            let totalRowIndex = -1;
            let totalRow = null;
            for (let i = 0; i < rows.length; i++) {
                const firstCell = String(rows[i]._rawData[1] || '').toLowerCase().trim();
                if (firstCell === 'total') {
                    totalRowIndex = i;
                    totalRow = rows[i]._rawData;
                    console.log('💰 Total 행 발견 (index:', i, ')');
                    break;
                }
            }

            if (!totalRow) {
                throw new Error('Total 행을 찾을 수 없습니다.');
            }

            // 3. 각 월별 데이터 초기화 (지출 먼저)
            const monthlyData = {};
            Object.keys(monthColumns).forEach(monthCode => {
                const colIndex = monthColumns[monthCode];

                // 지출은 total 행의 해당 월 열에서
                const expense = Math.abs(this._parseNumber(totalRow[colIndex]));

                monthlyData[monthCode] = {
                    month: this._getMonthName(monthCode),
                    income: 0, // 수입은 나중에 합산
                    expense: expense || 0
                };
            });

            // 4. Business Total 및 기타 지출 추가
            let businessTotalIndex = -1;
            for (let i = totalRowIndex + 1; i < rows.length; i++) {
                const cell0 = String(rows[i]._rawData[0] || '').toLowerCase().trim();
                const cell1 = String(rows[i]._rawData[1] || '').toLowerCase().trim();

                if (cell1 === 'total' && i > totalRowIndex) {
                    businessTotalIndex = i;
                    const businessRow = rows[i]._rawData;
                    console.log('💼 Business Total 행 발견 (index:', i, ')');

                    // Business 지출 합산
                    Object.keys(monthColumns).forEach(monthCode => {
                        const colIndex = monthColumns[monthCode];
                        const expense = Math.abs(this._parseNumber(businessRow[colIndex]));
                        if (expense > 0) {
                            monthlyData[monthCode].expense += expense;
                            console.log(`  📊 ${monthCode}월 Business 지출: ₩${expense.toLocaleString()}`);
                        }
                    });
                    break;
                }
            }

            // 5. Timeline 섹션 찾기
            let timelineIndex = -1;
            for (let i = 0; i < rows.length; i++) {
                const cell0 = String(rows[i]._rawData[0] || '').toLowerCase().trim();
                const cell1 = String(rows[i]._rawData[1] || '').toLowerCase().trim();

                if (cell0 === 'timeline' || cell1 === 'timeline') {
                    timelineIndex = i;
                    console.log('📍 Timeline 섹션 발견 (index:', i, ')');
                    break;
                }
            }

            // 6. Timeline Total 행에서 수입/지출 합산
            // Timeline 구조: B열=날짜, D열=항목, E열=지출, F열=수입, G열=잔고
            // 월별로 3칸씩: 1월(E,F,G=4,5,6), 2월(H,I,J=7,8,9), 3월(K,L,M=10,11,12)
            if (timelineIndex > 0) {
                console.log(`📊 Timeline Total 행 찾기...`);

                // Timeline 섹션의 마지막 Total 행 찾기
                for (let i = rows.length - 1; i > timelineIndex; i--) {
                    const cell0 = String(rows[i]._rawData[0] || '').toLowerCase().trim();
                    const cell1 = String(rows[i]._rawData[1] || '').toLowerCase().trim();

                    if (cell0 === 'total' || cell1 === 'total') {
                        const timelineTotalRow = rows[i]._rawData;
                        console.log('📊 Timeline Total 행 발견 (index:', i, ')');

                        // 각 월별로 수입 추출
                        Object.keys(monthColumns).forEach(monthCode => {
                            const monthColIndex = monthColumns[monthCode];

                            // Timeline에서 월 시작 열은 Loan/Business와 같지만
                            // E열=지출, F열=수입, G열=잔고 구조
                            // 1월: col 4(E)=지출, col 5(F)=수입
                            // 2월: col 7(H)=지출, col 8(I)=수입
                            // 3월: col 10(K)=지출, col 11(L)=수입

                            const incomeValue = this._parseNumber(timelineTotalRow[monthColIndex + 1]);

                            if (incomeValue > 0) {
                                monthlyData[monthCode].income = incomeValue;
                                console.log(`  💵 ${monthCode}월 Timeline 수입: ₩${incomeValue.toLocaleString()}`);
                            }
                        });

                        break;
                    }
                }
            } else {
                console.log('⚠️ Timeline 섹션을 찾을 수 없습니다.');
            }

            // 5. 결과 출력
            Object.keys(monthlyData).forEach(monthCode => {
                const data = monthlyData[monthCode];
                console.log(`✅ ${monthCode}월: 수입=₩${data.income.toLocaleString()}, 지출=₩${data.expense.toLocaleString()}`);
            });

            console.log(`✅ ${Object.keys(monthlyData).length}개월 데이터 조회 성공`);
            return monthlyData;
        } catch (error) {
            console.error('❌ 월별 데이터 조회 실패:', error);
            throw error;
        }
    }

    /**
     * 특정 월의 데이터 조회
     */
    async getMonthData(monthCode) {
        try {
            const allData = await this.getMonthlyData();

            if (!allData[monthCode]) {
                throw new Error(`${monthCode}월 데이터를 찾을 수 없습니다.`);
            }

            return allData[monthCode];
        } catch (error) {
            console.error(`❌ ${monthCode}월 데이터 조회 실패:`, error.message);
            throw error;
        }
    }

    /**
     * 오늘 현재 잔고 가져오기 (Timeline에서)
     */
    async getTodayBalance() {
        try {
            await this.initialize();

            const sheet = this.doc.sheetsByIndex[0];
            const rows = await sheet.getRows();

            // 오늘 날짜
            const today = new Date();
            const currentMonth = today.getMonth() + 1;
            const dayOfMonth = today.getDate();

            console.log(`📅 오늘 날짜: ${currentMonth}월 ${dayOfMonth}일`);

            // 월 인덱스 찾기
            await sheet.loadCells('A1:AN1');
            let monthColIndex = -1;
            for (let i = 0; i < sheet.columnCount; i++) {
                const cell = sheet.getCell(0, i);
                if (cell.value) {
                    const match = String(cell.value).match(/^(\d{2})/);
                    if (match && parseInt(match[1]) === currentMonth) {
                        monthColIndex = i;
                        break;
                    }
                }
            }

            if (monthColIndex === -1) {
                throw new Error(`${currentMonth}월 데이터를 찾을 수 없습니다.`);
            }

            // Timeline 찾기
            let timelineIndex = -1;
            for (let i = 0; i < rows.length; i++) {
                const cell0 = String(rows[i]._rawData[0] || '').toLowerCase().trim();
                if (cell0 === 'timeline') {
                    timelineIndex = i;
                    break;
                }
            }

            if (timelineIndex === -1) {
                throw new Error('Timeline 섹션을 찾을 수 없습니다.');
            }

            // Timeline에서 오늘 날짜 행 찾기
            for (let i = timelineIndex + 1; i < rows.length; i++) {
                const row = rows[i]._rawData;
                const dateCell = String(row[1] || '').trim();

                // 날짜가 일치하면
                if (dateCell === String(dayOfMonth)) {
                    // G열 = 잔고 (monthColIndex + 2)
                    const balance = this._parseNumber(row[monthColIndex + 2]);
                    console.log(`✅ ${currentMonth}월 ${dayOfMonth}일 잔고: ₩${balance.toLocaleString()}`);
                    return balance;
                }
            }

            throw new Error(`${currentMonth}월 ${dayOfMonth}일 데이터를 찾을 수 없습니다.`);
        } catch (error) {
            console.error('❌ 오늘 잔고 조회 실패:', error.message);
            throw error;
        }
    }

    /**
     * 기준날짜/기준금액 및 항목 리스트 가져오기 (B70:F107)
     */
    async getBaseDataAndItems() {
        try {
            await this.initialize();

            const sheet = this.doc.sheetsByIndex[0];

            // B70:B71 읽기 (기준날짜, 기준금액)
            await sheet.loadCells('B70:C70');
            const baseDateCell = sheet.getCell(69, 1); // B70 (0-indexed)
            const baseAmountCell = sheet.getCell(69, 2); // C70

            const baseDate = baseDateCell.value ? new Date(baseDateCell.value) : new Date();
            const baseAmount = this._parseNumber(baseAmountCell.value);

            console.log(`📅 기준날짜: ${baseDate.toLocaleDateString('ko-KR')}`);
            console.log(`💰 기준금액: ₩${baseAmount.toLocaleString()}`);

            // C95:E96 읽기 (카드 이용한도 및 현재 지출)
            await sheet.loadCells('C95:E96');
            const hyundaiLimit = this._parseNumber(sheet.getCell(94, 2).value); // C95
            const hyundaiUsed = this._parseNumber(sheet.getCell(94, 4).value); // E95
            const samsungLimit = this._parseNumber(sheet.getCell(95, 2).value); // C96
            const samsungUsed = this._parseNumber(sheet.getCell(95, 4).value); // E96

            console.log(`💳 현대카드: 한도 ₩${hyundaiLimit.toLocaleString()}, 사용 ₩${hyundaiUsed.toLocaleString()}`);
            console.log(`💳 삼성카드: 한도 ₩${samsungLimit.toLocaleString()}, 사용 ₩${samsungUsed.toLocaleString()}`);

            // A73:F107 읽기 (항목 리스트 - 결제수단 포함)
            await sheet.loadCells('A73:F107');
            const items = [];

            for (let row = 72; row < 107; row++) { // A73부터 F107까지 (0-indexed: 72~106)
                const paymentTypeCell = sheet.getCell(row, 0); // A열 = 결제수단 (현금/삼성카드/현대카드)
                const frequencyCell = sheet.getCell(row, 1); // B열 = 발생 주기
                const referenceCell = sheet.getCell(row, 2); // C열 = 참고 정보
                const itemNameCell = sheet.getCell(row, 3); // D열 = 항목명
                const expenseCell = sheet.getCell(row, 4); // E열 = 월별 지출액
                const incomeCell = sheet.getCell(row, 5); // F열 = 월별 수입액

                const paymentType = String(paymentTypeCell.value || '').trim();
                const frequency = String(frequencyCell.value || '').trim();
                const reference = String(referenceCell.value || '').trim();
                const itemName = String(itemNameCell.value || '').trim();
                const expense = this._parseNumber(expenseCell.value);
                const income = this._parseNumber(incomeCell.value);

                // 항목명이 있으면 추가
                if (itemName) {
                    items.push({
                        paymentType: paymentType, // 현금 / 삼성카드 / 현대카드
                        frequency: frequency, // 매월 / 매주 수요일 / 8월 1일 등
                        reference: reference,
                        itemName: itemName,
                        expense: expense,
                        income: income
                    });
                }
            }

            console.log(`✅ 항목 ${items.length}건 조회 성공`);
            return {
                baseDate: baseDate,
                baseAmount: baseAmount,
                items: items,
                cardInfo: {
                    hyundai: {
                        limit: hyundaiLimit,
                        used: hyundaiUsed,
                        remaining: hyundaiLimit - hyundaiUsed
                    },
                    samsung: {
                        limit: samsungLimit,
                        used: samsungUsed,
                        remaining: samsungLimit - samsungUsed
                    }
                }
            };
        } catch (error) {
            console.error('❌ 기준 데이터 조회 실패:', error.message);
            throw error;
        }
    }

    /**
     * 일일 기록 추가 (선택)
     */
    async recordDailyBalance(date, balance) {
        try {
            await this.initialize();

            // 두 번째 시트 (일일 기록) - 있으면
            if (this.doc.sheetsByIndex[1]) {
                const sheet = this.doc.sheetsByIndex[1];

                await sheet.addRow({
                    '날짜': new Date(date).toLocaleDateString('ko-KR'),
                    '잔고': balance,
                    '기록시간': new Date().toISOString()
                });

                console.log(`✅ ${date} 잔고 기록: ₩${balance.toLocaleString('ko-KR')}`);
            }
        } catch (error) {
            console.error('❌ 일일 기록 저장 실패:', error.message);
            // 에러가 나도 계속 진행
        }
    }

    /**
     * 오늘 잔고 계산 (기준날짜/금액 + 항목별 누적)
     */
    async calculateTodayBalance() {
        try {
            const data = await this.getBaseDataAndItems();
            const { baseDate, baseAmount, items } = data;

            const today = new Date();
            console.log(`📊 ${baseDate.toLocaleDateString('ko-KR')}부터 ${today.toLocaleDateString('ko-KR')}까지 계산`);

            let totalIncome = 0;
            let totalExpense = 0;

            items.forEach(item => {
                const count = this._calculateOccurrences(item.frequency, baseDate, today);
                 
                if (count > 0) {
                    const itemIncome = item.income * count;
                    const itemExpense = item.expense * count;

                    totalIncome += itemIncome;
                    totalExpense += itemExpense;

                    console.log(`  📌 ${item.itemName}: ${count}회 발생 (수입: +₩${itemIncome.toLocaleString()}, 지출: -₩${itemExpense.toLocaleString()})`);
                }
            });

            const currentBalance = baseAmount + totalIncome - totalExpense;

            console.log(`✅ 오늘 잔고 계산 완료:`);
            console.log(`   기준금액: ₩${baseAmount.toLocaleString()}`);
            console.log(`   총 수입: +₩${totalIncome.toLocaleString()}`);
            console.log(`   총 지출: -₩${totalExpense.toLocaleString()}`);
            console.log(`   현재 잔고: ₩${currentBalance.toLocaleString()}`);

            return currentBalance;
        } catch (error) {
            console.error('❌ 오늘 잔고 계산 실패:', error.message);
            throw error;
        }
    }

    /**
     * 발생 주기에 따라 기준날짜부터 오늘까지 몇 번 발생했는지 계산
     */
    _calculateOccurrences(frequency, baseDate, today) {
        const freq = frequency.toLowerCase().trim();
        // 1. 매월
        if (freq.startsWith('매월')) {
            const dayMatch = freq.match(/매월\s*(\d+)일/);
            if (!dayMatch) return 0;

            const targetDay = parseInt(dayMatch[1], 10);

            const year = today.getFullYear();
            const month = today.getMonth();

            const occurrenceDate = new Date(year, month, targetDay);

            // 기준일 이후 + 오늘 이전이면 1회
            if (occurrenceDate >= baseDate && occurrenceDate <= today) {
                return 1;
            }
            return 0;
        }

        // 2. 매주 (예: "매주 수요일")
        if (freq.includes('매주')) {
            const dayMatch = freq.match(/(월|화|수|목|금|토|일)/);
            if (dayMatch) {
                const targetDay = this._getDayOfWeek(dayMatch[1]);
                return this._countWeekdays(baseDate, today, targetDay);
            }
        }

        // 3. 특정일 (예: "8월 1일", "31일")
        const dateMatch = freq.match(/(\d+)월\s*(\d+)일/) || freq.match(/^(\d+)일$/);
        if (dateMatch) {
            if (dateMatch[2]) {
                // "8월 1일" 형식
                const targetMonth = parseInt(dateMatch[1]);
                const targetDay = parseInt(dateMatch[2]);
                return this._checkSpecificDate(baseDate, today, targetMonth, targetDay);
            } else {
                // "31일" 형식 (매월 31일)
                const targetDay = parseInt(dateMatch[1]);
                return this._countMonthlyDay(baseDate, today, targetDay);
            }
        }

        // 4. 기타 (1회만 발생)
        return 0;
    }

    /**
     * 요일 문자를 숫자로 변환 (일=0, 월=1, ..., 토=6)
     */
    _getDayOfWeek(dayStr) {
        const days = { '일': 0, '월': 1, '화': 2, '수': 3, '목': 4, '금': 5, '토': 6 };
        return days[dayStr] !== undefined ? days[dayStr] : -1;
    }

    /**
     * 기준날짜부터 오늘까지 특정 요일이 몇 번 나왔는지 계산
     */
    _countWeekdays(baseDate, today, targetDay) {
        let count = 0;
        const current = new Date(baseDate);

        while (current <= today) {
            if (current.getDay() === targetDay) {
                count++;
            }
            current.setDate(current.getDate() + 1);
        }

        return count;
    }

    /**
     * 특정 날짜(예: 8월 1일)가 기준날짜와 오늘 사이에 있는지 확인
     */
    _checkSpecificDate(baseDate, today, targetMonth, targetDay) {
        const targetDate = new Date(today.getFullYear(), targetMonth - 1, targetDay);

        // 기준날짜 이후이고 오늘 이전이면 1회
        if (targetDate >= baseDate && targetDate <= today) {
            return 1;
        }

        return 0;
    }

    /**
     * 매월 특정일(예: 31일)이 기준날짜부터 오늘까지 몇 번 발생했는지 계산
     */
    _countMonthlyDay(baseDate, today, targetDay) {
        let count = 0;
        const current = new Date(baseDate.getFullYear(), baseDate.getMonth(), 1);

        while (current <= today) {
            const daysInMonth = new Date(current.getFullYear(), current.getMonth() + 1, 0).getDate();

            // 해당 월에 targetDay가 존재하는지 확인
            if (targetDay <= daysInMonth) {
                const targetDate = new Date(current.getFullYear(), current.getMonth(), targetDay);

                // 기준날짜 이후이고 오늘 이전이면 카운트
                if (targetDate >= baseDate && targetDate <= today) {
                    count++;
                }
            }

            // 다음 달로 이동
            current.setMonth(current.getMonth() + 1);
        }

        return count;
    }

    /**
     * 숫자 파싱 헬퍼
     */
    _parseNumber(value) {
        if (!value) return 0;

        // 음수 기호와 쉼표 제거
        const cleaned = String(value)
            .replace(/[₩$¥]/g, '')
            .replace(/,/g, '')
            .replace(/[^0-9.-]/g, '');

        return parseInt(cleaned) || 0;
    }

    /**
     * 월 이름 변환
     */
    _getMonthName(monthCode) {
        const monthNum = parseInt(monthCode);
        if (monthNum >= 1 && monthNum <= 12) {
            return `${monthNum}월`;
        }
        return '미정';
    }

    /**
     * 초기화 상태 확인
     */
    isInitialized() {
        return this.initialized;
    }

    /**
     * 연결 리셋 (테스트용)
     */
    reset() {
        this.doc = null;
        this.initialized = false;
    }
}

module.exports = new GoogleSheetsService();