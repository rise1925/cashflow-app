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

            // 81행(Total) 로드: E81:AM81
            await sheet.loadCells('E81:AM81');

            const monthlyData = {};

            // 1~12월 반복
            for (let month = 1; month <= 12; month++) {
                const monthCode = String(month).padStart(2, '0');
                const cols = getMonthColumns(month);

                // 81행 (0-indexed: 80)
                const expenseCell = sheet.getCell(80, cols.expense);
                const incomeCell = sheet.getCell(80, cols.income);

                const monthExpense = Math.abs(this._parseNumber(expenseCell.value));
                const monthIncome = this._parseNumber(incomeCell.value);

                monthlyData[monthCode] = {
                    month: this._getMonthName(monthCode),
                    income: monthIncome,
                    expense: monthExpense
                };

                console.log(`✅ ${monthCode}월 Total(81행): 수입=₩${monthIncome.toLocaleString()}, 지출=₩${monthExpense.toLocaleString()}`);
            }

            // 여행 수익 추가 (B15:16, F/I/L.../AM 15:16)
            await sheet.loadCells('B15:AM16');

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
     * 특정 월의 상세 내역 조회 (36~80행)
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

            // A36:AN80 범위 로드 (날짜, 항목, 지출, 수입, 잔고)
            await sheet.loadCells('B36:AN80');

            const details = [];

            // 36~80행 순회 (0-indexed: 35~79)
            for (let row = 35; row < 80; row++) {
                const dateCell = sheet.getCell(row, 1); // A열 = 날짜
                const categoryCell = sheet.getCell(row, 3); // C열 = 카테고리 (현금/카드)
                const itemCell = sheet.getCell(row, 4); // D열 = 항목명
                const expenseCell = sheet.getCell(row, cols.expense);
                const incomeCell = sheet.getCell(row, cols.income);
                const balanceCell = sheet.getCell(row, cols.balance);

                const date = dateCell.value ? String(dateCell.value).trim() : '';
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
     * 예: 1월 22일 → G22 (1월 잔고 열의 22행)
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

            // 해당 날짜의 잔고 셀 로드 (예: G22)
            // 36행부터 시작이므로 실제 날짜 행 = 35 + dayOfMonth
            const rowIndex = 35 + dayOfMonth; // 0-indexed

            await sheet.loadCells(`A${rowIndex + 1}:AN${rowIndex + 1}`);

            const balanceCell = sheet.getCell(rowIndex, cols.balance);
            const balance = this._parseNumber(balanceCell.value);

            console.log(`✅ ${currentMonth}월 ${dayOfMonth}일 잔고 (행 ${rowIndex + 1}, 열 ${cols.balance}): ₩${balance.toLocaleString()}`);
            return balance;
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

            // C18:G19 읽기 (카드 정보)
            // C18=삼성카드 한도, D18=이름, G18=남은 한도
            // C19=현대카드 한도, D19=이름, G19=남은 한도
            await sheet.loadCells('C18:G19');
            const samsungLimit = this._parseNumber(sheet.getCell(17, 2).value); // C18
            const samsungName = String(sheet.getCell(17, 3).value || '삼성카드').trim(); // D18
            const samsungRemaining = this._parseNumber(sheet.getCell(17, 6).value); // G18
            const samsungUsed = samsungLimit - samsungRemaining;

            const hyundaiLimit = this._parseNumber(sheet.getCell(18, 2).value); // C19
            const hyundaiName = String(sheet.getCell(18, 3).value || '현대카드').trim(); // D19
            const hyundaiRemaining = this._parseNumber(sheet.getCell(18, 6).value); // G19
            const hyundaiUsed = hyundaiLimit - hyundaiRemaining;

            console.log(`💳 ${samsungName}: 한도 ₩${samsungLimit.toLocaleString()}, 남은 ₩${samsungRemaining.toLocaleString()}, 사용 ₩${samsungUsed.toLocaleString()}`);
            console.log(`💳 ${hyundaiName}: 한도 ₩${hyundaiLimit.toLocaleString()}, 남은 ₩${hyundaiRemaining.toLocaleString()}, 사용 ₩${hyundaiUsed.toLocaleString()}`);

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