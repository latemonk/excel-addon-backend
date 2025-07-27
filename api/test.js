// 간단한 테스트 엔드포인트
export default function handler(req, res) {
  // 모든 origin 허용
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  // OPTIONS 요청 처리
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }
  
  // 응답
  res.status(200).json({
    message: 'Test endpoint is working!',
    method: req.method,
    timestamp: new Date().toISOString(),
    headers: req.headers
  });
}