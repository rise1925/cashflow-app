const { GoogleSpreadsheet } = require('google-spreadsheet');
const config = require('../../config');

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

            // 설정 확인
            if (!config.GOOGLE_SHEET_ID) {
                throw new Error('GOOGLE_SHEET_ID가 설정되지 않았습니다.');
            }
            if (!config.GOOGLE_SERVICE_ACCOUNT_EMAIL) {
                throw new Error('GOOGLE_SERVICE_ACCOUNT_EMAIL이 설정되지 않았습니다.');
            }
            if (!config.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY) {
                throw new Error('GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY가 설정되지 않았습니다.');
            }

            console.log('🔑 Google Sheets 인증 시작...');
            console.log('  Sheet ID:', config.GOOGLE_SHEET_ID);
            console.log('  Email:', config.GOOGLE_SERVICE_ACCOUNT_EMAIL);

            // google-spreadsheet 3.x 방식
            this.doc = new GoogleSpreadsheet(config.GOOGLE_SHEET_ID);

            // Private Key 처리 (개행 문자 복원)
            const privateKey = config.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY.replace(/\\n/g, '\n');

            // 서비스 계정으로 인증
            await this.doc.useServiceAccountAuth({
                client_email: config.GOOGLE_SERVICE_ACCOUNT_EMAIL,
                private_key: privateKey
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
     * 월별 데이터 조회 (새 구조)
     * 1월 Total: E81=지출, F81=수입
     * 2월 Total: H81=지출, I81=수입
     * ...
     * 12월 Total: AL81=지출, AM81=수입
     */
    async getMonthlyData() {
        try {
            await this.initialize();

            const sheet = this.doc.sheetsByIndex[0];
            console.log('📄 시트 이름:', sheet.title);

            // 열 매핑: 1월=E/F/G (col 4/5/6), 2월=H/I/J (col 7/8/9), ...
            // 패턴: 월마다 3칸씩 이동
            const getMonthColumns = (monthNum) => {
                const baseCol = 4 + (monthNum - 1) * 3; // 1월=4, 2월=7, 3월=10, ...
                return {
                    expense: baseCol,      // E, H, K, ...
                    income: baseCol + 1,   // F, I, L, ...
                    balance: baseCol + 2   // G, J, M, ...
                };
            };

            // 80~81행 로드: E80:AN81 (80행=월말잔고, 81행=Total)
            // 12월까지 포함: E(4), F(5), G(6) ... AL(37-expense), AM(38-income), AN(39-balance)
            await sheet.loadCells('E80:AN81');

            const monthlyData = {};

            // 1~12월 반복
            for (let month = 1; month <= 12; month++) {
                const monthCode = String(month).padStart(2, '0');
                const cols = getMonthColumns(month);

                // 81행 (0-indexed: 80) - Total
                const expenseCell = sheet.getCell(80, cols.expense);
                const incomeCell = sheet.getCell(80, cols.income);

                // 80행 (0-indexed: 79) - 월말 잔고
                const balanceCell = sheet.getCell(79, cols.balance);

                const monthExpense = Math.abs(this._parseNumber(expenseCell.value));
                const monthIncome = this._parseNumber(incomeCell.value);
                const monthEndBalance = this._parseNumber(balanceCell.value);

                monthlyData[monthCode] = {
                    month: this._getMonthName(monthCode),
                    income: monthIncome,
                    expense: monthExpense,
                    balance: monthEndBalance
                };

                console.log(`✅ ${monthCode}월: 수입=₩${monthIncome.toLocaleString()}, 지출=₩${monthExpense.toLocaleString()}, 월말잔고=₩${monthEndBalance.toLocaleString()}`);
            }

            // 여행 수익 추가 (B15:16, F/I/L.../AN 15:16)
            await sheet.loadCells('B15:AN16');

            console.log('🌏 여행 수익 데이터 로딩...');

            // 15행, 16행에서 날짜 확인
            for (let rowIdx = 14; rowIdx <= 15; rowIdx++) { // 0-indexed: 14, 15
                const dateCell = sheet.getCell(rowIdx, 1); // B열
                const itemCell = sheet.getCell(rowIdx, 3); // D열

                console.log(`  행 ${rowIdx + 1}: 날짜=${dateCell.value}, 항목=${itemCell.value}`);

                // 1~12월 각각의 수익 열 확인
                for (let month = 1; month <= 12; month++) {
                    const monthCode = String(month).padStart(2, '0');
                    const cols = getMonthColumns(month);

                    // 여행 수익은 income 열(F, I, L, ...)에서 가져오기
                    const travelIncomeCell = sheet.getCell(rowIdx, cols.income);
                    const travelIncome = this._parseNumber(travelIncomeCell.value);

                    if (travelIncome > 0) {
                        monthlyData[monthCode].income += travelIncome;
                        console.log(`    💰 ${monthCode}월 여행 수익 추가: ₩${travelIncome.toLocaleString()}`);
                    }
                }
            }

            console.log(`✅ ${Object.keys(monthlyData).length}개월 데이터 조회 완료 (여행 수익 포함)`);
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
     * 특정 월의 상세 내역 조회 (26~34행: 특수항목, 36~80행: 일일 데이터)
     * @param {number} month - 월 번호 (1~12)
     * @returns {Promise<Array>} 상세 내역 배열
     */
    async getMonthDetailData(month) {
        try {
            await this.initialize();

            const sheet = this.doc.sheetsByIndex[0];
            console.log(`📊 ${month}월 상세 내역 조회 중...`);

            // 열 매핑
            const getMonthColumns = (monthNum) => {
                const baseCol = 4 + (monthNum - 1) * 3;
                return {
                    expense: baseCol,      // E, H, K, ...
                    income: baseCol + 1,   // F, I, L, ...
                    balance: baseCol + 2   // G, J, M, ...
                };
            };

            const cols = getMonthColumns(month);

            // B26:AN80 범위 로드 (26~34행: 특수항목, 36~80행: 일일 데이터)
            await sheet.loadCells('B26:AN80');

            const details = [];

            // 26~34행 순회 (특수 항목들, 0-indexed: 25~33)
            for (let row = 25; row < 34; row++) {
                const dateCell = sheet.getCell(row, 1); // B열 = 날짜
                const categoryCell = sheet.getCell(row, 2); // C열 = 카테고리
                const itemCell = sheet.getCell(row, 3); // D열 = 항목명
                const expenseCell = sheet.getCell(row, cols.expense);
                const incomeCell = sheet.getCell(row, cols.income);
                const balanceCell = sheet.getCell(row, cols.balance);

                // 날짜 파싱 (Excel 날짜, 텍스트, 숫자 처리)
                const date = this._parseDateValue(dateCell.value);
                const category = categoryCell.value ? String(categoryCell.value).trim() : '';
                const item = itemCell.value ? String(itemCell.value).trim() : '';
                const expense = Math.abs(this._parseNumber(expenseCell.value));
                const income = this._parseNumber(incomeCell.value);
                const balance = this._parseNumber(balanceCell.value);

                // 지출이나 수입이 있으면 추가 (특수 항목은 날짜/항목명이 없을 수 있음)
                if (expense > 0 || income > 0 || item) {
                    details.push({
                        date: date,
                        category: category,
                        item: item,
                        expense: expense,
                        income: income,
                        balance: balance
                    });
                }
            }

            // 36~80행 순회 (일일 데이터, 0-indexed: 35~79)
            for (let row = 35; row < 80; row++) {
                const dateCell = sheet.getCell(row, 1); // B열 = 날짜
                const categoryCell = sheet.getCell(row, 2); // C열 = 카테고리 (현금/카드)
                const itemCell = sheet.getCell(row, 3); // D열 = 항목명
                const expenseCell = sheet.getCell(row, cols.expense);
                const incomeCell = sheet.getCell(row, cols.income);
                const balanceCell = sheet.getCell(row, cols.balance);

                // 날짜 파싱 (Excel 날짜, 텍스트, 숫자 처리)
                const date = this._parseDateValue(dateCell.value);
                const category = categoryCell.value ? String(categoryCell.value).trim() : '';
                const item = itemCell.value ? String(itemCell.value).trim() : '';
                const expense = Math.abs(this._parseNumber(expenseCell.value));
                const income = this._parseNumber(incomeCell.value);
                const balance = this._parseNumber(balanceCell.value);

                // 날짜나 항목이 있으면 추가
                if (date || item || expense > 0 || income > 0) {
                    details.push({
                        date: date,
                        category: category,
                        item: item,
                        expense: expense,
                        income: income,
                        balance: balance
                    });
                }
            }

            console.log(`✅ ${month}월 상세 내역 ${details.length}건 조회 완료`);
            return details;
        } catch (error) {
            console.error(`❌ ${month}월 상세 내역 조회 실패:`, error.message);
            throw error;
        }
    }

    /**
     * 오늘 현재 잔고 가져오기
     * B36:B80 범위에서 오늘 날짜를 찾아서 해당 행의 잔고 반환
     * 예: 1월 22일 → B70에서 22를 찾아 G70의 잔고 반환
     */
    async getTodayBalance() {
        try {
            await this.initialize();

            const sheet = this.doc.sheetsByIndex[0];

            // 오늘 날짜
            const today = new Date();
            const currentMonth = today.getMonth() + 1;
            const dayOfMonth = today.getDate();

            console.log(`📅 오늘 날짜: ${currentMonth}월 ${dayOfMonth}일`);

            // 열 매핑: 1월=G (col 6), 2월=J (col 9), ...
            const getMonthColumns = (monthNum) => {
                const baseCol = 4 + (monthNum - 1) * 3;
                return {
                    expense: baseCol,      // E, H, K, ...
                    income: baseCol + 1,   // F, I, L, ...
                    balance: baseCol + 2   // G, J, M, ...
                };
            };

            const cols = getMonthColumns(currentMonth);

            // B36:AN80 범위 로드 (B열=날짜)
            await sheet.loadCells('B36:AN80');

            // 36~80행에서 B열의 날짜가 오늘 날짜와 일치하는 행 찾기
            for (let row = 35; row < 80; row++) {
                const dateCell = sheet.getCell(row, 1);
                const rawValue = dateCell.value;

                let cellDay = null;

                if (rawValue instanceof Date) {
                    cellDay = rawValue.getDate();
                } else {
                    const parsed = parseInt(String(rawValue).replace(/[^0-9]/g, ''), 10);
                    cellDay = isNaN(parsed) ? null : parsed;
                }
                 
                if (cellDay === dayOfMonth) {
                    const balanceCell = sheet.getCell(row, cols.balance);
                    const balance = this._parseNumber(balanceCell.value);

                    console.log(`✅ 잔고: ₩${balance.toLocaleString()}`);
                    return balance;
                }
            }


            // 날짜를 찾지 못한 경우
            console.warn(`⚠️ ${currentMonth}월 ${dayOfMonth}일 데이터를 찾을 수 없습니다.`);
            return 0;
        } catch (error) {
            console.error('❌ 오늘 잔고 조회 실패:', error.message);
            throw error;
        }
    }

    /**
     * 이번달 월말 예상 잔고 가져오기 (80번째 행)
     * 예: 1월이면 G80, 2월이면 J80
     */
    async getMonthEndBalance() {
        try {
            await this.initialize();

            const sheet = this.doc.sheetsByIndex[0];

            // 현재 월
            const today = new Date();
            const currentMonth = today.getMonth() + 1;

            console.log(`📅 현재 월: ${currentMonth}월`);

            // 열 매핑
            const getMonthColumns = (monthNum) => {
                const baseCol = 4 + (monthNum - 1) * 3;
                return {
                    expense: baseCol,
                    income: baseCol + 1,
                    balance: baseCol + 2
                };
            };

            const cols = getMonthColumns(currentMonth);

            // 80행 로드 (0-indexed: 79)
            await sheet.loadCells('E80:AN80');

            const balanceCell = sheet.getCell(79, cols.balance);
            const monthEndBalance = this._parseNumber(balanceCell.value);

            console.log(`✅ ${currentMonth}월 월말 예상 잔고 (80행): ₩${monthEndBalance.toLocaleString()}`);
            return monthEndBalance;
        } catch (error) {
            console.error('❌ 월말 예상 잔고 조회 실패:', error.message);
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
     * 날짜 값 파싱 헬퍼 (Excel 날짜 시리얼 번호, Date 객체, 텍스트 처리)
     */
    _parseDateValue(value) {
        if (!value) return '';

        // 1. Date 객체인 경우
        if (value instanceof Date) {
            return value.getDate(); // 일(day)만 반환
        }

        // 2. 숫자인 경우 (Excel 날짜 시리얼 번호일 가능성)
        if (typeof value === 'number') {
            // Excel 날짜는 1900년 1월 1일 기준 일수
            // 46113 같은 큰 숫자는 날짜 시리얼 번호
            if (value > 1000) {
                const excelEpoch = new Date(1900, 0, 1);
                const days = value - 2; // Excel 버그 보정
                const date = new Date(excelEpoch.getTime() + days * 24 * 60 * 60 * 1000);
                return date.getDate(); // 일(day)만 반환
            }
            // 작은 숫자는 그대로 반환 (1~31)
            return value;
        }

        // 3. 문자열인 경우
        const str = String(value).trim();

        // "1일", "22일" 형식에서 숫자만 추출
        const dayMatch = str.match(/^(\d+)일?$/);
        if (dayMatch) {
            return parseInt(dayMatch[1], 10);
        }

        // "매주 수요일" 같은 텍스트는 그대로 반환
        return str;
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