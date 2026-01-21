const express = require('express');
const router = express.Router();
const googleSheets = require('../services/googleSheets');
const { 
    calculateBalance, 
    generateProjection,
    generateMonthComparison
} = require('../utils/calculator');

/**
 * GET /api/health
 * API 헬스 체크
 */
router.get('/health', (req, res) => {
    res.json({
        status: 'OK',
        service: 'Cashflow API',
        timestamp: new Date().toISOString()
    });
});

/**
 * GET /api/monthly-data
 * Google Sheets에서 월별 데이터 조회
 */
router.get('/monthly-data', async (req, res) => {
    try {
        const data = await googleSheets.getMonthlyData();
        
        res.json({
            success: true,
            message: '월별 데이터 조회 성공',
            data: data,
            count: Object.keys(data).length
        });
    } catch (error) {
        console.error('API 에러:', error);
        res.status(500).json({
            success: false,
            error: 'Google Sheets 데이터 조회 실패',
            message: error.message
        });
    }
});

/**
 * GET /api/monthly-data/:month
 * 특정 월의 데이터 조회
 */
router.get('/monthly-data/:month', async (req, res) => {
    try {
        const { month } = req.params;
        const data = await googleSheets.getMonthData(month);

        res.json({
            success: true,
            message: `${month}월 데이터 조회 성공`,
            data: data
        });
    } catch (error) {
        console.error('API 에러:', error);
        res.status(400).json({
            success: false,
            error: '월별 데이터 조회 실패',
            message: error.message
        });
    }
});

/**
 * GET /api/today-balance
 * Google Sheets Timeline에서 오늘 날짜의 실제 잔고 조회
 */
router.get('/today-balance', async (req, res) => {
    try {
        const balance = await googleSheets.getTodayBalance();

        res.json({
            success: true,
            message: '오늘 잔고 조회 성공',
            data: {
                balance: balance,
                date: new Date().toISOString()
            }
        });
    } catch (error) {
        console.error('API 에러:', error);
        res.status(500).json({
            success: false,
            error: '오늘 잔고 조회 실패',
            message: error.message
        });
    }
});

/**
 * GET /api/calculated-balance
 * 기준날짜/금액 기반으로 오늘 잔고 계산
 */
router.get('/calculated-balance', async (req, res) => {
    try {
        const balance = await googleSheets.calculateTodayBalance();

        res.json({
            success: true,
            message: '오늘 잔고 계산 성공',
            data: {
                balance: balance,
                date: new Date().toISOString()
            }
        });
    } catch (error) {
        console.error('API 에러:', error);
        res.status(500).json({
            success: false,
            error: '오늘 잔고 계산 실패',
            message: error.message
        });
    }
});

/**
 * GET /api/items
 * 기준날짜/금액 및 항목 리스트 조회
 */
router.get('/items', async (req, res) => {
    try {
        const data = await googleSheets.getBaseDataAndItems();

        res.json({
            success: true,
            message: '항목 데이터 조회 성공',
            data: data
        });
    } catch (error) {
        console.error('API 에러:', error);
        res.status(500).json({
            success: false,
            error: '항목 데이터 조회 실패',
            message: error.message
        });
    }
});

/**
 * POST /api/calculate
 * 현재 잔고 계산
 * 
 * Request Body:
 * {
 *   "monthlyData": { ... },
 *   "startBalance": 2337459,
 *   "date": "2025-02-01"
 * }
 */
router.post('/calculate', (req, res) => {
    try {
        const { monthlyData, startBalance, date } = req.body;

        // 입력값 검증
        if (!monthlyData || typeof monthlyData !== 'object') {
            return res.status(400).json({
                success: false,
                error: '유효한 monthlyData가 필요합니다.'
            });
        }

        if (!startBalance || typeof startBalance !== 'number') {
            return res.status(400).json({
                success: false,
                error: '유효한 startBalance가 필요합니다.'
            });
        }

        const today = date ? new Date(date) : new Date();
        
        if (isNaN(today.getTime())) {
            return res.status(400).json({
                success: false,
                error: '유효한 날짜 형식이 아닙니다.'
            });
        }

        const result = calculateBalance(startBalance, monthlyData, today);
        
        res.json({
            success: true,
            message: '잔고 계산 완료',
            data: result,
            calculatedAt: new Date().toISOString()
        });
    } catch (error) {
        console.error('API 에러:', error);
        res.status(500).json({
            success: false,
            error: '잔고 계산 실패',
            message: error.message
        });
    }
});

/**
 * POST /api/projection
 * 30일 현금흐름 예측
 * 
 * Request Body:
 * {
 *   "currentBalance": 1500000,
 *   "dailyExpense": 173025,
 *   "daysForward": 30
 * }
 */
router.post('/projection', (req, res) => {
    try {
        const { currentBalance, dailyExpense, daysForward = 30 } = req.body;

        if (!currentBalance || typeof currentBalance !== 'number') {
            return res.status(400).json({
                success: false,
                error: '유효한 currentBalance가 필요합니다.'
            });
        }

        if (!dailyExpense || typeof dailyExpense !== 'number') {
            return res.status(400).json({
                success: false,
                error: '유효한 dailyExpense가 필요합니다.'
            });
        }

        const projection = generateProjection(currentBalance, dailyExpense, daysForward);

        res.json({
            success: true,
            message: '현금흐름 예측 완료',
            data: projection
        });
    } catch (error) {
        console.error('API 에러:', error);
        res.status(500).json({
            success: false,
            error: '예측 생성 실패',
            message: error.message
        });
    }
});

/**
 * POST /api/comparison
 * 월별 비교 데이터
 * 
 * Request Body:
 * {
 *   "monthlyData": { ... }
 * }
 */
router.post('/comparison', (req, res) => {
    try {
        const { monthlyData } = req.body;

        if (!monthlyData || typeof monthlyData !== 'object') {
            return res.status(400).json({
                success: false,
                error: '유효한 monthlyData가 필요합니다.'
            });
        }

        const comparison = generateMonthComparison(monthlyData);

        res.json({
            success: true,
            message: '월별 비교 데이터 생성 완료',
            data: comparison
        });
    } catch (error) {
        console.error('API 에러:', error);
        res.status(500).json({
            success: false,
            error: '비교 데이터 생성 실패',
            message: error.message
        });
    }
});

/**
 * POST /api/record
 * 일일 잔고 기록 (선택)
 * 
 * Request Body:
 * {
 *   "date": "2025-02-01",
 *   "balance": 2337459
 * }
 */
router.post('/record', async (req, res) => {
    try {
        const { date, balance } = req.body;

        if (!date || !balance) {
            return res.status(400).json({
                success: false,
                error: '날짜와 잔고가 필요합니다.'
            });
        }

        await googleSheets.recordDailyBalance(date, balance);

        res.json({
            success: true,
            message: '잔고가 기록되었습니다.',
            data: { date, balance }
        });
    } catch (error) {
        console.error('API 에러:', error);
        res.status(500).json({
            success: false,
            error: '기록 저장 실패',
            message: error.message
        });
    }
});

/**
 * POST /api/full-analysis
 * 전체 분석 (월별 + 계산 + 예측)
 * 
 * Request Body:
 * {
 *   "startBalance": 2337459,
 *   "date": "2025-02-01"
 * }
 */
router.post('/full-analysis', async (req, res) => {
    try {
        const { startBalance, date } = req.body;

        if (!startBalance) {
            return res.status(400).json({
                success: false,
                error: '초기 잔고가 필요합니다.'
            });
        }

        // 월별 데이터 조회
        const monthlyData = await googleSheets.getMonthlyData();
        
        // 현재 잔고 계산
        const today = date ? new Date(date) : new Date();
        const balance = calculateBalance(startBalance, monthlyData, today);
        
        // 예측 생성
        const projection = generateProjection(
            balance.currentBalance,
            balance.metrics.dailyExpense,
            30
        );
        
        // 월별 비교
        const comparison = generateMonthComparison(monthlyData);

        res.json({
            success: true,
            message: '전체 분석 완료',
            data: {
                balance,
                projection,
                comparison,
                monthlyData
            },
            analyzedAt: new Date().toISOString()
        });
    } catch (error) {
        console.error('API 에러:', error);
        res.status(500).json({
            success: false,
            error: '전체 분석 실패',
            message: error.message
        });
    }
});

module.exports = router;