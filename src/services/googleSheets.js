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

            // 81~82행 로드: E81:AN82 (81행=월말잔고, 82행=Total) - 대출 행 추가로 +1
            // 12월까지 포함: E(4), F(5), G(6) ... AL(37-expense), AM(38-income), AN(39-balance)
            await sheet.loadCells('E81:AN82');

            const monthlyData = {};

            // 1~12월 반복
            for (let month = 1; month <= 12; month++) {
                const monthCode = String(month).padStart(2, '0');
                const cols = getMonthColumns(month);

                // 82행 (0-indexed: 81) - 지출, 수입 (Total)
                const expenseCell = sheet.getCell(81, cols.expense); // E82, H82, K82...
                const incomeCell = sheet.getCell(81, cols.income);   // F82, I82, L82...

                // 81행 (0-indexed: 80) - 월말 누적 잔고
                const balanceCell = sheet.getCell(80, cols.balance); // G81, J81, M81...

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

            // 여행 수익 추가 (B16:17, F/I/L.../AN 16:17) - 대출 행 추가로 +1
            await sheet.loadCells('B16:AN17');

            console.log('🌏 여행 수익 데이터 로딩...');

            // 16행, 17행에서 날짜 확인 (대출 행 추가로 +1)
            for (let rowIdx = 15; rowIdx <= 16; rowIdx++) { // 0-indexed: 15, 16
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

            // B27:AN81 범위 로드 (27~35행: 특수항목, 37~81행: 일일 데이터) - 대출 행 추가로 +1
            await sheet.loadCells('B27:AN81');

            const details = [];
            const specialItems = []; // 27~35행 특수 항목 수집

            // 27~35행 순회 (특수 항목들, 0-indexed: 26~34)
            for (let row = 26; row < 35; row++) {
                const dateCell = sheet.getCell(row, 1); // B열 = 발생 주기 (예: "매주 수요일")
                const categoryCell = sheet.getCell(row, 2); // C열 = 카테고리
                const itemCell = sheet.getCell(row, 3); // D열 = 항목명
                const expenseCell = sheet.getCell(row, cols.expense);
                const incomeCell = sheet.getCell(row, cols.income);

                const frequency = dateCell.value ? String(dateCell.value).trim() : '';
                const category = categoryCell.value ? String(categoryCell.value).trim() : '';
                const item = itemCell.value ? String(itemCell.value).trim() : '';
                const totalExpense = Math.abs(this._parseNumber(expenseCell.value));
                const totalIncome = this._parseNumber(incomeCell.value);

                // 지출이나 수입이 있고 발생 주기가 있으면 특수 항목으로 수집
                if ((totalExpense > 0 || totalIncome > 0) && frequency) {
                    specialItems.push({
                        frequency: frequency,
                        category: category,
                        item: item,
                        totalExpense: totalExpense,
                        totalIncome: totalIncome
                    });
                }
            }

            // 특수 항목을 실제 날짜로 분배
            const year = new Date().getFullYear();
            specialItems.forEach(specialItem => {
                const dates = this._getOccurrenceDates(specialItem.frequency, month, year);
                const occurrenceCount = dates.length;

                if (occurrenceCount > 0) {
                    const expensePerOccurrence = Math.round(specialItem.totalExpense / occurrenceCount);
                    const incomePerOccurrence = Math.round(specialItem.totalIncome / occurrenceCount);

                    dates.forEach(day => {
                        details.push({
                            date: day,
                            category: specialItem.category,
                            item: specialItem.item,
                            expense: expensePerOccurrence,
                            income: incomePerOccurrence,
                            balance: 0 // 잔고는 나중에 계산
                        });
                    });
                }
            });

            // 37~81행 순회 (일일 데이터, 0-indexed: 36~80) - 대출 행 추가로 +1
            for (let row = 36; row < 81; row++) {
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

            // B37:AN81 범위 로드 (B열=날짜) - 대출 행 추가로 +1
            await sheet.loadCells('B37:AN81');

            // 37~81행에서 B열의 날짜가 오늘 날짜와 일치하는 행 찾기 (0-indexed: 36~80)
            for (let row = 36; row < 81; row++) {
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

            // 81행 로드 (0-indexed: 80) - 대출 행 추가로 +1
            await sheet.loadCells('E81:AN81');

            const balanceCell = sheet.getCell(80, cols.balance);
            const monthEndBalance = this._parseNumber(balanceCell.value);

            console.log(`✅ ${currentMonth}월 월말 예상 잔고 (81행): ₩${monthEndBalance.toLocaleString()}`);
            return monthEndBalance;
        } catch (error) {
            console.error('❌ 월말 예상 잔고 조회 실패:', error.message);
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

            // B71:C71 읽기 (기준날짜, 기준금액) - 대출 행 추가로 +1
            await sheet.loadCells('B71:C71');
            const baseDateCell = sheet.getCell(70, 1); // B71 (0-indexed: 70)
            const baseAmountCell = sheet.getCell(70, 2); // C71

            const baseDate = baseDateCell.value ? new Date(baseDateCell.value) : new Date();
            const baseAmount = this._parseNumber(baseAmountCell.value);

            console.log(`📅 기준날짜: ${baseDate.toLocaleDateString('ko-KR')}`);
            console.log(`💰 기준금액: ₩${baseAmount.toLocaleString()}`);

            // C19:G20 읽기 (카드 정보) - 대출 행 추가로 +1
            // C19=삼성카드 한도, D19=이름, G19=남은 한도
            // C20=현대카드 한도, D20=이름, G20=남은 한도
            await sheet.loadCells('C19:G20');
            const samsungLimit = this._parseNumber(sheet.getCell(18, 2).value); // C19 (0-indexed: 18)
            const samsungName = String(sheet.getCell(18, 3).value || '삼성카드').trim(); // D19
            const samsungRemaining = this._parseNumber(sheet.getCell(18, 6).value); // G19
            const samsungUsed = samsungLimit - samsungRemaining;

            const hyundaiLimit = this._parseNumber(sheet.getCell(19, 2).value); // C20 (0-indexed: 19)
            const hyundaiName = String(sheet.getCell(19, 3).value || '현대카드').trim(); // D20
            const hyundaiRemaining = this._parseNumber(sheet.getCell(19, 6).value); // G20
            const hyundaiUsed = hyundaiLimit - hyundaiRemaining;

            console.log(`💳 ${samsungName}: 한도 ₩${samsungLimit.toLocaleString()}, 남은 ₩${samsungRemaining.toLocaleString()}, 사용 ₩${samsungUsed.toLocaleString()}`);
            console.log(`💳 ${hyundaiName}: 한도 ₩${hyundaiLimit.toLocaleString()}, 남은 ₩${hyundaiRemaining.toLocaleString()}, 사용 ₩${hyundaiUsed.toLocaleString()}`);

            // A74:F101 읽기 (항목 리스트 - 결제수단 포함) - 대출 행 추가로 +1
            await sheet.loadCells('A74:F101');
            const items = [];

            for (let row = 73; row < 101; row++) { // A74부터 F101까지 (0-indexed: 73~100)
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
     * 발생 주기에 따라 해당 월의 실제 날짜 배열 반환
     * @param {string} frequency - "매주 수요일", "매월 5일" 등
     * @param {number} month - 월 (1~12)
     * @param {number} year - 년도
     * @returns {Array<number>} - 날짜 배열 (예: [7, 14, 21, 28])
     */
    _getOccurrenceDates(frequency, month, year) {
        const freq = frequency.toLowerCase().trim();
        const dates = [];

        // 1. 매주 X요일
        if (freq.includes('매주')) {
            const dayNames = { '일': 0, '월': 1, '화': 2, '수': 3, '목': 4, '금': 5, '토': 6 };
            const dayMatch = freq.match(/(월|화|수|목|금|토|일)/);

            if (dayMatch) {
                const targetDayOfWeek = dayNames[dayMatch[1]];
                const monthStart = new Date(year, month - 1, 1);
                const monthEnd = new Date(year, month, 0);

                for (let d = new Date(monthStart); d <= monthEnd; d.setDate(d.getDate() + 1)) {
                    if (d.getDay() === targetDayOfWeek) {
                        dates.push(d.getDate());
                    }
                }
            }
        }
        // 2. 매월 X일
        else if (freq.includes('매월')) {
            const dayMatch = freq.match(/(\d+)일/);
            if (dayMatch) {
                const day = parseInt(dayMatch[1], 10);
                const monthEnd = new Date(year, month, 0).getDate();
                if (day <= monthEnd) {
                    dates.push(day);
                }
            }
        }
        // 3. 특정 월일 (예: "1월 24일")
        else if (freq.match(/(\d+)월\s*(\d+)일/)) {
            const specificDateMatch = freq.match(/(\d+)월\s*(\d+)일/);
            if (specificDateMatch) {
                const targetMonth = parseInt(specificDateMatch[1]);
                const targetDay = parseInt(specificDateMatch[2]);
                if (targetMonth === month) {
                    dates.push(targetDay);
                }
            }
        }
        // 4. 숫자만 (예: "25") - 매월 25일로 해석
        else if (freq.match(/^\d+$/)) {
            const day = parseInt(freq, 10);
            const monthEnd = new Date(year, month, 0).getDate();
            if (day <= monthEnd) {
                dates.push(day);
            }
        }

        return dates;
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
     * 대출 상환 데이터 조회 (B3:AN15)
     * B3~B15: 상환 날짜 (일자)
     * C3~C15: 대출 기관/채권자
     * D3~D15: 상환 방법/이자율
     * E/F/G (1월), H/I/J (2월), ... AN (12월)
     * E열: 월 상환액, F열: 대출 잔액, G열: 남은 상환 횟수
     */
    async getLoanData() {
        try {
            await this.initialize();

            const sheet = this.doc.sheetsByIndex[0];
            console.log('💳 대출 상환 데이터 조회 중...');

            // B3:AN15 범위 로드 (행 확장: 10 → 15)
            await sheet.loadCells('B3:AN15');

            const loans = [];

            // 3~15행 순회 (0-indexed: 2~14)
            for (let row = 2; row < 15; row++) {
                const dateCell = sheet.getCell(row, 1); // B열 = 상환일
                const lenderCell = sheet.getCell(row, 2); // C열 = 대출기관
                const methodCell = sheet.getCell(row, 3); // D열 = 상환방법/이자율

                const paymentDay = this._parseDateValue(dateCell.value);
                const lender = lenderCell.value ? String(lenderCell.value).trim() : '';
                const method = methodCell.value ? String(methodCell.value).trim() : '';

                // 대출 기관이 있으면 처리
                if (lender) {
                    const monthlyData = [];

                    // 1~12월 데이터 수집
                    for (let month = 1; month <= 12; month++) {
                        const baseCol = 4 + (month - 1) * 3;
                        const expenseCell = sheet.getCell(row, baseCol); // E열: 월 상환액 (지출)
                        const balanceCell = sheet.getCell(row, baseCol + 1); // F열: 대출 잔액
                        const monthsCell = sheet.getCell(row, baseCol + 2); // G열: 남은 상환 횟수

                        const payment = Math.abs(this._parseNumber(expenseCell.value));
                        const balance = this._parseNumber(balanceCell.value);
                        const months = this._parseNumber(monthsCell.value);

                        if (month === 1) {
                            console.log(`  ${lender} - ${month}월: 상환액=₩${payment.toLocaleString()}, 잔액=₩${balance.toLocaleString()}, 잔여회차=${months}회`);
                            console.log(`    셀값 원본 - E${row+1}(상환액): ${expenseCell.value}, F${row+1}(잔액): ${balanceCell.value}, G${row+1}(회차): ${monthsCell.value}`);
                        }

                        monthlyData.push({
                            month: month,
                            payment: payment, // 월 상환액
                            remainingBalance: balance, // 잔여 대출 잔액
                            remainingMonths: months // 남은 상환 횟수
                        });
                    }

                    loans.push({
                        paymentDay: paymentDay,
                        lender: lender,
                        method: method,
                        monthlyData: monthlyData
                    });

                    console.log(`✅ ${lender}: 상환일 ${paymentDay}일, ${method}`);
                }
            }

            console.log(`✅ 대출 ${loans.length}건 조회 완료`);
            return loans;
        } catch (error) {
            console.error('❌ 대출 데이터 조회 실패:', error.message);
            throw error;
        }
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