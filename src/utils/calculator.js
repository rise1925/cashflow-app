/**
 * 일일 메트릭 계산
 */
function calculateDailyMetrics(monthlyData, today) {
    const currentMonth = today.getMonth() + 1;
    const dayOfMonth = today.getDate();
    
    const monthCode = String(currentMonth).padStart(2, '0');
    const current = monthlyData[monthCode];
    
    if (!current) {
        return null;
    }

    // 일일 평균 지출 (월 지출 / 30)
    const dailyExpense = Math.round(current.expense / 30);
    
    // 오늘까지의 누적 지출
    const accumulatedExpense = dailyExpense * dayOfMonth;
    
    // 오늘 수입 (월급 입금일: 7, 14, 21, 28일)
    const salaryDays = [7, 14, 21, 28];
    const todayIncome = salaryDays.includes(dayOfMonth) ? current.income : 0;
    
    // 남은 일수
    const daysInMonth = 30;
    const daysRemaining = daysInMonth - dayOfMonth;

    return {
        monthCode,
        month: current.month,
        dailyExpense,
        accumulatedExpense,
        todayIncome,
        todayExpense: dailyExpense,
        daysInMonth,
        dayOfMonth,
        daysRemaining,
        monthlyIncome: current.income,
        monthlyExpense: current.expense
    };
}

/**
 * 현재 잔고 계산
 */
function calculateBalance(startBalance, monthlyData, today) {
    const metrics = calculateDailyMetrics(monthlyData, today);
    
    if (!metrics) {
        return null;
    }

    // 현재 잔고 계산
    const currentBalance = startBalance 
        - metrics.accumulatedExpense 
        + metrics.todayIncome;
    
    // 마이너스까지 남은 일수
    let daysToNegative = 0;
    if (currentBalance > 0) {
        daysToNegative = Math.ceil(currentBalance / metrics.dailyExpense);
    }

    // 월말 예상 잔고
    const projectedMonthEndBalance = currentBalance 
        - (metrics.dailyExpense * metrics.daysRemaining);

    // 상태 판단
    let status = 'safe';
    if (currentBalance < 0) {
        status = 'negative';
    } else if (daysToNegative <= 5) {
        status = 'critical';
    } else if (daysToNegative <= 30) {
        status = 'warning';
    }

    return {
        currentBalance: Math.round(currentBalance),
        daysToNegative,
        status,
        projectedMonthEndBalance: Math.round(projectedMonthEndBalance),
        metrics,
        details: {
            startBalance,
            accumulatedExpense: Math.round(metrics.accumulatedExpense),
            accumulatedIncome: metrics.todayIncome,
            remainingMonthDays: metrics.daysRemaining
        }
    };
}

/**
 * 30일 예측 데이터 생성
 */
function generateProjection(currentBalance, dailyExpense, daysForward = 30) {
    const days = [];
    const balances = [];
    let balance = currentBalance;

    for (let i = 0; i <= daysForward; i++) {
        days.push(i + '일');
        balances.push(Math.round(balance));
        balance -= dailyExpense;
    }

    return {
        days,
        balances,
        firstNegativeDay: balances.findIndex(b => b < 0)
    };
}

/**
 * 월별 비교 데이터
 */
function generateMonthComparison(monthlyData) {
    const comparison = {};

    Object.keys(monthlyData).forEach(monthCode => {
        const data = monthlyData[monthCode];
        const netFlow = data.income - data.expense;
        
        comparison[monthCode] = {
            month: data.month,
            income: data.income,
            expense: data.expense,
            netFlow,
            status: netFlow < 0 ? 'deficit' : 'surplus'
        };
    });

    return comparison;
}

/**
 * 잔고 상태 텍스트
 */
function getStatusText(status) {
    const statusMap = {
        'negative': '🚨 마이너스 발생',
        'critical': '⚠️ 긴급 상황',
        'warning': '⚠️ 주의 필요',
        'safe': '✅ 안정적'
    };
    return statusMap[status] || '상태 미정';
}

/**
 * 마이너스 예상 날짜 계산
 */
function calculateNegativeDate(today, daysToNegative) {
    const negativeDate = new Date(today);
    negativeDate.setDate(negativeDate.getDate() + daysToNegative);
    return negativeDate;
}

module.exports = {
    calculateDailyMetrics,
    calculateBalance,
    generateProjection,
    generateMonthComparison,
    getStatusText,
    calculateNegativeDate
};