// Upstash Redis를 사용한 인증키 관리 API
import { Redis } from '@upstash/redis';

// Redis 클라이언트 초기화
const redis = new Redis({
  url: process.env.UPSTASH_REDIS_REST_URL,
  token: process.env.UPSTASH_REDIS_REST_TOKEN,
});

// 관리자 인증 확인
function isAdmin(req) {
  const adminPassword = req.headers['x-admin-password'];
  return adminPassword === process.env.ADMIN_PASSWORD;
}

// 랜덤 키 생성
function generateAuthKey() {
  const prefix = 'WORKS';
  const timestamp = Date.now().toString(36).toUpperCase();
  const random = Math.random().toString(36).substr(2, 6).toUpperCase();
  return `${prefix}-${timestamp}-${random}`;
}

export default async function handler(req, res) {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Admin-Password');
  
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  // 관리자 인증 확인
  if (!isAdmin(req)) {
    return res.status(401).json({ error: '관리자 인증이 필요합니다.' });
  }

  try {
    switch (req.method) {
      case 'GET':
        // 모든 인증키 조회
        const keys = await redis.smembers('auth_keys') || [];
        const keyDetails = [];
        
        for (const key of keys) {
          const details = await redis.hgetall(`auth_key:${key}`);
          keyDetails.push({
            key,
            ...details
          });
        }
        
        return res.status(200).json({ 
          success: true,
          keys: keyDetails.sort((a, b) => b.createdAt - a.createdAt)
        });

      case 'POST':
        // 새 인증키 생성
        const { company, memo } = req.body;
        
        if (!company) {
          return res.status(400).json({ error: '회사명은 필수입니다.' });
        }
        
        const newKey = generateAuthKey();
        const keyData = {
          company,
          memo: memo || '',
          createdAt: new Date().toISOString(),
          createdBy: 'admin',
          isActive: true,
          usageCount: 0
        };
        
        // Redis에 저장
        await redis.sadd('auth_keys', newKey);
        await redis.hset(`auth_key:${newKey}`, keyData);
        
        return res.status(200).json({ 
          success: true,
          key: newKey,
          ...keyData
        });

      case 'DELETE':
        // 인증키 삭제/비활성화
        const { key } = req.body;
        
        if (!key) {
          return res.status(400).json({ error: '삭제할 키를 지정해주세요.' });
        }
        
        // 완전 삭제 대신 비활성화
        await redis.hset(`auth_key:${key}`, { isActive: false });
        
        return res.status(200).json({ 
          success: true,
          message: '인증키가 비활성화되었습니다.'
        });

      default:
        return res.status(405).json({ error: 'Method not allowed' });
    }
  } catch (error) {
    console.error('Auth key management error:', error);
    return res.status(500).json({ 
      error: 'Redis 연결 오류가 발생했습니다. Upstash Redis가 설정되었는지 확인하세요.',
      details: error.message
    });
  }
}