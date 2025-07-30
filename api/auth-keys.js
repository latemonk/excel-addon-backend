// Upstash Redis를 사용한 인증키 관리 API
import { Redis } from '@upstash/redis';

let redis = null;
// In-memory storage fallback
const memoryStorage = {
  keys: new Set(),
  data: new Map()
};

// Redis 클라이언트 초기화 시도
try {
  if (process.env.UPSTASH_REDIS_REST_URL && process.env.UPSTASH_REDIS_REST_TOKEN) {
    redis = new Redis({
      url: process.env.UPSTASH_REDIS_REST_URL,
      token: process.env.UPSTASH_REDIS_REST_TOKEN,
    });
  }
} catch (error) {
  console.log('Redis initialization failed, using in-memory storage:', error);
}

// 관리자 인증 확인
function isAdmin(req) {
  const adminPassword = req.headers['x-admin-password'];
  return adminPassword === process.env.ADMIN_PASSWORD;
}

// 랜덤 키 생성
function generateAuthKey() {
  const prefix = 'WRKS';
  // 8자리 랜덤 문자열 생성 (대문자 + 숫자)
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let randomPart = '';
  for (let i = 0; i < 8; i++) {
    randomPart += characters.charAt(Math.floor(Math.random() * characters.length));
  }
  return `${prefix}-${randomPart}`;
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
        let keys = [];
        const keyDetails = [];
        
        if (redis) {
          keys = await redis.smembers('auth_keys') || [];
          for (const key of keys) {
            const details = await redis.hgetall(`auth_key:${key}`);
            keyDetails.push({
              key,
              ...details
            });
          }
        } else {
          // In-memory fallback
          keys = Array.from(memoryStorage.keys);
          for (const key of keys) {
            const details = memoryStorage.data.get(key);
            if (details) {
              keyDetails.push({
                key,
                ...details
              });
            }
          }
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
          isActive: 'true',  // Redis에 문자열로 저장
          usageCount: 0
        };
        
        // 저장
        if (redis) {
          await redis.sadd('auth_keys', newKey);
          await redis.hset(`auth_key:${newKey}`, keyData);
        } else {
          // In-memory fallback
          memoryStorage.keys.add(newKey);
          memoryStorage.data.set(newKey, keyData);
        }
        
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
        if (redis) {
          await redis.hset(`auth_key:${key}`, { isActive: 'false' });  // 문자열로 저장
        } else {
          // In-memory fallback
          const keyData = memoryStorage.data.get(key);
          if (keyData) {
            keyData.isActive = 'false';  // 문자열로 저장
            memoryStorage.data.set(key, keyData);
          }
        }
        
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
      error: '서버 오류가 발생했습니다.',
      details: error.message,
      storageMode: redis ? 'redis' : 'memory'
    });
  }
}