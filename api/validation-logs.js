// Validation logs API for viewing auth key usage history
let redis = null;

// Redis 클라이언트 초기화 시도
try {
  if (process.env.UPSTASH_REDIS_REST_URL && process.env.UPSTASH_REDIS_REST_TOKEN) {
    const { Redis } = await import('@upstash/redis');
    redis = new Redis({
      url: process.env.UPSTASH_REDIS_REST_URL,
      token: process.env.UPSTASH_REDIS_REST_TOKEN,
    });
  }
} catch (error) {
  console.log('Redis initialization failed:', error);
}

// 관리자 인증 확인
function isAdmin(req) {
  const adminPassword = req.headers['x-admin-password'];
  return adminPassword === process.env.ADMIN_PASSWORD;
}

export default async function handler(req, res) {
  // CORS headers - Use environment variable
  const origin = req.headers.origin;
  const ALLOWED_ORIGINS = process.env.ALLOWED_ORIGINS?.split(',').map(origin => origin.trim()) || [
    'https://localhost:3000',
    'https://excel.office.com',
    'https://latemonk.github.io'
  ];
  
  // Set CORS headers
  if (origin && ALLOWED_ORIGINS.includes(origin)) {
    res.setHeader('Access-Control-Allow-Origin', origin);
  } else {
    res.setHeader('Access-Control-Allow-Origin', '*');
  }
  
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Admin-Password');
  
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  // 관리자 인증 확인
  if (!isAdmin(req)) {
    return res.status(401).json({ error: '관리자 인증이 필요합니다.' });
  }

  if (req.method !== 'GET') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const logs = [];
    
    if (redis) {
      // Get all log keys
      const logKeys = await redis.smembers('validation_logs') || [];
      
      // Get details for each log
      for (const logKey of logKeys) {
        const logData = await redis.hgetall(logKey);
        if (logData) {
          logs.push(logData);
        }
      }
      
      // Sort by timestamp (newest first)
      logs.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
      
      // Limit to recent 100 logs
      const recentLogs = logs.slice(0, 100);
      
      return res.status(200).json({ 
        success: true,
        logs: recentLogs,
        total: logs.length
      });
    } else {
      // No Redis available
      return res.status(200).json({ 
        success: true,
        logs: [],
        total: 0,
        message: 'Redis not configured - logs not available'
      });
    }
  } catch (error) {
    console.error('Validation logs error:', error);
    return res.status(500).json({ 
      error: '로그 조회 중 오류가 발생했습니다.',
      details: error.message
    });
  }
}