const express = require('express');
const cors = require('cors');
require('dotenv').config();
const apiRoutes = require('./src/routes/api');

const app = express();
const PORT = process.env.PORT || 8001;

// 미들웨어
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('src/public'));

// API 라우트
app.use('/api', apiRoutes);

// 메인 페이지
app.get('/', (req, res) => {
    res.sendFile(__dirname + '/src/public/index.html');
});

// 헬스 체크
app.get('/health', (req, res) => {
    res.json({ 
        status: 'OK',
        timestamp: new Date().toISOString(),
        uptime: process.uptime()
    });
});

// 404 핸들러
app.use((req, res) => {
    res.status(404).json({ 
        error: 'Not Found',
        message: `${req.method} ${req.path} 경로를 찾을 수 없습니다.`
    });
});

// 에러 핸들러
app.use((err, req, res, next) => {
    console.error('❌ 에러:', err);
    res.status(err.status || 500).json({ 
        error: '서버 오류가 발생했습니다.',
        message: process.env.NODE_ENV === 'development' ? err.message : '오류 정보'
    });
});

app.listen(PORT, () => {
    console.log(`🚀 서버가 포트 ${PORT}에서 실행 중입니다.`);
    console.log(`📱 http://localhost:${PORT} 접속하세요.`);
    console.log(`🏥 헬스체크: http://localhost:${PORT}/health`);
});

module.exports = app;