# Node.js 현금흐름 앱 - 개발부터 배포까지 완전 가이드

## 📋 목차
1. [프로젝트 생성](#1-프로젝트-생성)
2. [Google Sheets API 연동](#2-google-sheets-api-연동)
3. [로컬 개발 환경](#3-로컬-개발-환경)
4. [카페24 호스팅 배포](#4-카페24-호스팅-배포)

---

## 1. 프로젝트 생성

### 1-1. 프로젝트 폴더 생성

```bash
# 프로젝트 폴더 생성
mkdir cashflow-app
cd cashflow-app

# Node.js 프로젝트 초기화
npm init -y
```

### 1-2. 필요한 패키지 설치

```bash
# 필수 패키지
npm install express cors dotenv axios
npm install --save-dev nodemon

# Google Sheets API
npm install google-auth-library google-spreadsheet
```

### 1-3. 프로젝트 구조

```
cashflow-app/
├── .env                    # 환경변수 (보안정보)
├── .gitignore              # Git 무시 파일
├── package.json            # 프로젝트 설정
├── server.js               # Express 메인 서버
├── public/                 # 프론트엔드 (HTML/CSS/JS)
│   ├── index.html
│   ├── css/
│   │   └── style.css
│   └── js/
│       └── app.js
├── src/                    # 백엔드 코드
│   ├── routes/
│   │   └── api.js
│   ├── services/
│   │   └── googleSheets.js
│   └── utils/
│       └── calculator.js
└── credentials.json        # Google API 키 (절대 Git에 올리면 안됨!)
```

---

## 2. Google Sheets API 연동

### 2-1. Google Cloud 프로젝트 생성

#### Step 1: Google Cloud Console 접속
1. https://console.cloud.google.com 방문
2. 새 프로젝트 생성 → "cashflow-app" 이름으로 생성

#### Step 2: API 활성화
1. API 및 서비스 > 라이브러리 검색
2. "Google Sheets API" 검색 후 활성화
3. "Google Drive API" 검색 후 활성화

#### Step 3: 서비스 계정 생성
1. API 및 서비스 > 사용자 인증 정보
2. "사용자 인증 정보 만들기" > "서비스 계정"
3. 서비스 계정명: "cashflow-service" 입력
4. "만들기 및 계속" 클릭

#### Step 4: 키 생성
1. 생성된 서비스 계정 클릭
2. "키" 탭 > "새 키 추가" > "JSON"
3. **credentials.json 파일이 다운로드됨** (매우 중요!)
4. 프로젝트 폴더에 저장

### 2-2. Google Sheets 공유 설정

1. 당신의 Google Sheets 열기
2. 공유 버튼 클릭
3. credentials.json의 "client_email" 값을 복사하여 공유 상대로 추가
4. 편집자 권한으로 공유

### 2-3. Sheets ID 확인

```
https://docs.google.com/spreadsheets/d/[SHEET_ID]/edit
                                        ^^^^^^^^^^^^^^^^
                                        이 부분이 Sheet ID
```

---

## 3. 로컬 개발 환경

### 3-1. 환경변수 설정 (.env)

```env
# 포트
PORT=3000

# Google Sheets
GOOGLE_SHEET_ID=당신의_스프레드시트_ID_입력
GOOGLE_CREDENTIALS_PATH=./credentials.json

# 환경
NODE_ENV=development
```

### 3-2. 서버 코드 (server.js)

```javascript
const express = require('express');
const cors = require('cors');
require('dotenv').config();
const apiRoutes = require('./src/routes/api');

const app = express();
const PORT = process.env.PORT || 3000;

// 미들웨어
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// API 라우트
app.use('/api', apiRoutes);

// 메인 페이지
app.get('/', (req, res) => {
    res.sendFile(__dirname + '/public/index.html');
});

// 에러 핸들링
app.use((err, req, res, next) => {
    console.error('에러:', err);
    res.status(500).json({ 
        error: '서버 오류가 발생했습니다.',
        message: err.message 
    });
});

app.listen(PORT, () => {
    console.log(`🚀 서버가 ${PORT}번 포트에서 실행 중입니다.`);
    console.log(`📱 http://localhost:${PORT} 접속하세요.`);
});

module.exports = app;
```

### 3-3. Google Sheets 서비스 (src/services/googleSheets.js)

```javascript
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

class GoogleSheetsService {
    constructor() {
        this.doc = null;
    }

    async initialize() {
        try {
            const credentials = require('../../credentials.json');
            
            const auth = new JWT({
                email: credentials.client_email,
                key: credentials.private_key,
                scopes: ['https://www.googleapis.com/auth/spreadsheets']
            });

            this.doc = new GoogleSpreadsheet(
                process.env.GOOGLE_SHEET_ID,
                auth
            );

            await this.doc.loadInfo();
            console.log('✅ Google Sheets 연동 성공');
        } catch (error) {
            console.error('❌ Google Sheets 연동 실패:', error.message);
            throw error;
        }
    }

    async getMonthlyData() {
        try {
            if (!this.doc) await this.initialize();

            // 첫 번째 시트 가져오기 (대출 정보)
            const sheet = this.doc.sheetsByIndex[0];
            const rows = await sheet.getRows();

            const monthlyData = {};
            
            // 헤더 파싱
            const headers = rows[0]._rawData;
            
            // 데이터 추출 (월별)
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                const month = row._rawData[0]; // 월 번호
                
                if (month && month.includes('01')) {
                    monthlyData['01'] = {
                        month: '1월',
                        income: parseInt(row._rawData[3]) || 0,
                        expense: parseInt(row._rawData[4]) || 0
                    };
                }
                // ... 다른 월들도 유사하게 처리
            }

            return monthlyData;
        } catch (error) {
            console.error('데이터 조회 실패:', error.message);
            throw error;
        }
    }

    async updateBalance(date, balance) {
        try {
            if (!this.doc) await this.initialize();
            
            const sheet = this.doc.sheetsByIndex[1]; // 두 번째 시트 (일일 데이터)
            
            await sheet.addRow({
                'Date': date,
                'Balance': balance,
                'UpdatedAt': new Date().toISOString()
            });

            console.log(`✅ ${date} 잔고 업데이트: ₩${balance}`);
        } catch (error) {
            console.error('업데이트 실패:', error.message);
            throw error;
        }
    }
}

module.exports = new GoogleSheetsService();
```

### 3-4. 계산 유틸 (src/utils/calculator.js)

```javascript
function calculateDailyMetrics(monthlyData, today) {
    const currentMonth = today.getMonth() + 1;
    const dayOfMonth = today.getDate();
    
    const current = monthlyData[String(currentMonth).padStart(2, '0')];
    
    if (!current) {
        return null;
    }

    // 일일 평균 지출
    const dailyExpense = current.expense / 30;
    
    // 오늘까지의 누적 지출
    const accumulatedExpense = dailyExpense * dayOfMonth;
    
    // 오늘 수입
    const salaryDays = [7, 14, 21, 28];
    const todayIncome = salaryDays.includes(dayOfMonth) ? current.income : 0;
    
    return {
        dailyExpense: Math.round(dailyExpense),
        accumulatedExpense: Math.round(accumulatedExpense),
        todayIncome,
        daysInMonth: 30,
        currentMonth
    };
}

function calculateBalance(startBalance, monthlyData, today) {
    const metrics = calculateDailyMetrics(monthlyData, today);
    if (!metrics) return null;

    const currentBalance = startBalance - metrics.accumulatedExpense + metrics.todayIncome;
    const daysToNegative = currentBalance > 0 
        ? Math.ceil(currentBalance / metrics.dailyExpense) 
        : 0;

    return {
        currentBalance,
        daysToNegative,
        metrics
    };
}

module.exports = {
    calculateDailyMetrics,
    calculateBalance
};
```

### 3-5. API 라우트 (src/routes/api.js)

```javascript
const express = require('express');
const router = express.Router();
const googleSheets = require('../services/googleSheets');
const { calculateBalance } = require('../utils/calculator');

// Google Sheets 데이터 조회
router.get('/monthly-data', async (req, res) => {
    try {
        const data = await googleSheets.getMonthlyData();
        res.json({
            success: true,
            data: data
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// 현재 잔고 계산
router.post('/calculate-balance', (req, res) => {
    try {
        const { monthlyData, startBalance, date } = req.body;
        const today = new Date(date);
        
        const result = calculateBalance(startBalance, monthlyData, today);
        
        res.json({
            success: true,
            data: result
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// 잔고 업데이트
router.post('/update-balance', async (req, res) => {
    try {
        const { date, balance } = req.body;
        await googleSheets.updateBalance(date, balance);
        
        res.json({
            success: true,
            message: '잔고가 업데이트되었습니다.'
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

module.exports = router;
```

### 3-6. 프론트엔드 (public/index.html)

```html
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>현금흐름 앱</title>
    <link rel="stylesheet" href="/css/style.css">
</head>
<body>
    <div id="app">
        <!-- 동적으로 로드될 컨텐츠 -->
    </div>
    <script src="/js/app.js"></script>
</body>
</html>
```

### 3-7. package.json 스크립트 설정

```json
{
  "name": "cashflow-app",
  "version": "1.0.0",
  "description": "일일 현금흐름 추적 앱",
  "main": "server.js",
  "scripts": {
    "start": "node server.js",
    "dev": "nodemon server.js",
    "test": "echo \"Error: no test specified\" && exit 1"
  },
  "keywords": [],
  "author": "",
  "license": "ISC",
  "dependencies": {
    "axios": "^1.6.0",
    "cors": "^2.8.5",
    "dotenv": "^16.3.1",
    "express": "^4.18.2",
    "google-auth-library": "^9.2.0",
    "google-spreadsheet": "^4.1.1"
  },
  "devDependencies": {
    "nodemon": "^3.0.2"
  }
}
```

### 3-8. .gitignore

```
node_modules/
.env
credentials.json
.DS_Store
*.log
dist/
build/
```

### 3-9. 로컬 실행

```bash
# 개발 모드 (자동 재시작)
npm run dev

# 프로덕션 모드
npm start

# http://localhost:3000 접속
```

---

## 4. 카페24 호스팅 배포

### 4-1. 카페24 가입 및 호스팅 구매

1. https://www.cafe24.com 접속
2. 회원가입 및 로그인
3. 호스팅 > Node.js 호스팅 선택
   - 패키지: 기본 또는 스탠다드
   - 기간: 1개월 또는 연간 선택
4. 결제 완료

### 4-2. 카페24 FTP/SSH 접속 정보 확인

1. 호스팅 관리자 페이지 로그인
2. "호스팅 관리" > "서비스 관리"
3. 다음 정보 복사:
   - FTP 주소
   - FTP 계정명
   - FTP 비밀번호
   - SSH 정보 (있으면)

### 4-3. 카페24에 배포

#### 방법 1: FTP 사용 (쉬움)

```bash
# 1. 카페24 계정 정보로 FTP 접속
ftp ftp주소
# 계정명, 비밀번호 입력

# 2. 공개 폴더로 이동
cd public_html

# 3. 프로젝트 전체 업로드 (node_modules, credentials.json 제외)
put server.js
put package.json
put .env    # (또는 카페24 대시보드에서 환경변수 설정)

mkdir src
mkdir public
# ... src, public 폴더 안의 파일들도 업로드
```

#### 방법 2: SSH 사용 (추천)

```bash
# 1. SSH로 접속
ssh 카페24_계정@호스트명

# 2. 공개 폴더로 이동
cd public_html

# 3. Git 클론 (있으면)
git clone https://github.com/당신의계정/cashflow-app.git
cd cashflow-app

# 4. npm 의존성 설치
npm install --production

# 5. PM2로 앱 실행 (백그라운드 유지)
npm install -g pm2
pm2 start server.js --name "cashflow-app"
pm2 startup
pm2 save

# 6. 상태 확인
pm2 status
```

### 4-4. 환경변수 설정

#### 방법 1: SSH에서 .env 파일 생성

```bash
nano .env
```

```env
PORT=3000
GOOGLE_SHEET_ID=당신의_스프레드시트_ID
GOOGLE_CREDENTIALS_PATH=./credentials.json
NODE_ENV=production
```

Save: `Ctrl + X` → `Y` → `Enter`

#### 방법 2: 카페24 대시보드에서 설정

1. 호스팅 관리 > 환경설정
2. Node.js 환경변수에 추가

### 4-5. 카페24 도메인 설정

1. 호스팅 관리 > 도메인 관리
2. 보유 도메인 또는 새 도메인 연결
3. DNS 설정:
   - A 레코드: 카페24 서버 IP
   - 또는 CNAME: 카페24 제공 URL

### 4-6. HTTPS 설정 (무료 SSL)

1. 호스팅 관리 > 보안설정
2. "Let's Encrypt SSL 인증서" > 설치
3. 자동 갱신 활성화

### 4-7. PM2를 이용한 자동 재시작

```bash
# 카페24 SSH에서

# PM2 설치
npm install -g pm2

# 앱 시작
pm2 start server.js --name "cashflow-app"

# 부팅 시 자동 시작
pm2 startup
pm2 save

# 로그 보기
pm2 logs cashflow-app

# 앱 재시작
pm2 restart cashflow-app

# 앱 중지
pm2 stop cashflow-app
```

### 4-8. 배포 후 확인

```bash
# 1. 배포된 사이트 접속
https://당신의도메인.com

# 2. API 테스트
curl https://당신의도메인.com/api/monthly-data

# 3. 로그 확인
pm2 logs cashflow-app

# 4. 문제 발생 시
pm2 stop cashflow-app
npm start  # 오류 메시지 확인
```

---

## 5. Google Sheets 데이터 형식

### 첫 번째 시트: 월별 대출 정보

| 월 | 이름 | 월급 | 지출 | 대출1 | 대출2 |
|----|------|------|------|-------|-------|
| 01 | 1월 | 1500000 | 5206003 | 400000 | 300000 |
| 02 | 2월 | 0 | 5206003 | 400000 | 300000 |
| 03 | 3월 | 3800000 | 5206003 | 400000 | 300000 |

### 두 번째 시트: 일일 기록

| Date | Balance | UpdatedAt |
|------|---------|-----------|
| 2025-02-01 | 2337459 | 2025-02-01T09:00:00Z |

---

## 6. 문제 해결

### 에러: "credentials.json을 찾을 수 없음"
```
→ credentials.json이 프로젝트 루트에 있는지 확인
→ .env의 GOOGLE_CREDENTIALS_PATH가 올바른지 확인
```

### 에러: "Google Sheets에 접근할 수 없음"
```
→ 서비스 계정 이메일을 Sheets에 공유했는지 확인
→ Google Drive API가 활성화되어 있는지 확인
```

### 카페24에서 앱이 시작되지 않음
```
→ PM2 로그 확인: pm2 logs
→ Node.js 버전 확인: node --version
→ npm install 재실행
```

### HTTPS 연결 안 됨
```
→ SSL 인증서가 설치되었는지 확인
→ 도메인 DNS 레코드 확인
→ 24시간 대기 (전파 시간)
```

---

## 7. 배포 체크리스트

- [ ] Google Cloud 프로젝트 생성
- [ ] Google Sheets API 활성화
- [ ] 서비스 계정 생성 및 credentials.json 다운로드
- [ ] Google Sheets 공유 설정
- [ ] 로컬에서 `npm run dev` 테스트
- [ ] 카페24 호스팅 구매
- [ ] 카페24에 파일 업로드
- [ ] PM2로 앱 시작
- [ ] 도메인 연결
- [ ] HTTPS 설정
- [ ] https://당신의도메인.com 접속 확인

---

## 8. 유용한 명령어

```bash
# 로컬 개발
npm run dev                    # 자동 재시작으로 개발
npm start                      # 프로덕션 시작

# PM2 (카페24)
pm2 start server.js            # 앱 시작
pm2 stop cashflow-app         # 앱 중지
pm2 restart cashflow-app      # 앱 재시작
pm2 logs cashflow-app         # 실시간 로그
pm2 delete cashflow-app       # 앱 삭제
pm2 list                       # 실행 중인 앱 목록

# Git (선택)
git init                       # Git 초기화
git add .                      # 파일 추가
git commit -m "메시지"         # 커밋
git push origin main           # GitHub 푸시

# Node.js
npm install package-name       # 패키지 설치
npm update                     # 패키지 업데이트
npm list                       # 설치된 패키지 목록
```

---

## 📞 도움말

- **Google Cloud 문서**: https://cloud.google.com/docs
- **Express.js 문서**: https://expressjs.com
- **PM2 문서**: https://pm2.keymetrics.io
- **카페24 지원**: https://www.cafe24.com/support

---

**마지막 팁**: 
- 항상 `.env`와 `credentials.json`을 Git에 올리지 마세요 (보안!)
- 배포 전에 로컬에서 철저히 테스트하세요
- 카페24 대시보드에서 리소스 사용량을 정기적으로 확인하세요
- PM2 로그를 통해 에러를 빠르게 파악할 수 있습니다