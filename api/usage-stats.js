// 회사별 월별 사용자 통계 API
import { Redis } from '@upstash/redis';

let redis = null;

// Redis 클라이언트 초기화 시도
try {
  if (process.env.UPSTASH_REDIS_REST_URL && process.env.UPSTASH_REDIS_REST_TOKEN) {
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
    const stats = {
      companies: {},
      totalUniqueUsers: 0,
      totalFreeUsers: 0,
      totalPaidUsers: 0,
      currentMonth: new Date().toISOString().substring(0, 7), // YYYY-MM
      breakdown: {
        free: {
          totalUsers: 0,
          currentMonthUsers: 0,
          monthlyActiveUsers: {}
        },
        paid: {
          totalUsers: 0,
          currentMonthUsers: 0,
          monthlyActiveUsers: {}
        }
      }
    };
    
    if (redis) {
      // Get all log keys
      const logKeys = await redis.smembers('validation_logs') || [];
      
      // Process logs to extract unique users per company per month
      const companyMonthlyUsers = {};
      const freeUsersByMonth = {};
      const paidUsersByMonth = {};
      const allFreeUsers = new Set();
      const allPaidUsers = new Set();
      
      for (const logKey of logKeys) {
        const logData = await redis.hgetall(logKey);
        if (logData && logData.email && logData.company && logData.timestamp) {
          const month = logData.timestamp.substring(0, 7); // YYYY-MM
          const email = logData.email.toLowerCase();
          const company = logData.company;
          const isFreeUser = logData.isFreeUser === 'true' || logData.isFreeUser === true || 
                             logData.authKey === 'Free' || company === 'Free User';
          
          // Initialize company data if not exists
          if (!companyMonthlyUsers[company]) {
            companyMonthlyUsers[company] = {};
          }
          
          // Initialize month data if not exists
          if (!companyMonthlyUsers[company][month]) {
            companyMonthlyUsers[company][month] = new Set();
          }
          
          // Add unique user
          companyMonthlyUsers[company][month].add(email);
          
          // Track free vs paid users
          if (isFreeUser) {
            allFreeUsers.add(email);
            if (!freeUsersByMonth[month]) {
              freeUsersByMonth[month] = new Set();
            }
            freeUsersByMonth[month].add(email);
          } else {
            allPaidUsers.add(email);
            if (!paidUsersByMonth[month]) {
              paidUsersByMonth[month] = new Set();
            }
            paidUsersByMonth[month].add(email);
          }
        }
      }
      
      // Convert Sets to counts
      for (const [company, monthlyData] of Object.entries(companyMonthlyUsers)) {
        const isFreeCompany = company === 'Free User';
        stats.companies[company] = {
          monthlyActiveUsers: {},
          totalUniqueUsers: new Set(),
          currentMonthUsers: 0,
          isFree: isFreeCompany
        };
        
        for (const [month, userSet] of Object.entries(monthlyData)) {
          stats.companies[company].monthlyActiveUsers[month] = userSet.size;
          
          // Add all users to total unique users
          userSet.forEach(user => stats.companies[company].totalUniqueUsers.add(user));
          
          // Count current month users
          if (month === stats.currentMonth) {
            stats.companies[company].currentMonthUsers = userSet.size;
          }
        }
        
        // Convert total unique users Set to count
        const totalCount = stats.companies[company].totalUniqueUsers.size;
        stats.companies[company].totalUniqueUsers = totalCount;
        stats.totalUniqueUsers += totalCount;
      }
      
      // Process free/paid user breakdown
      stats.totalFreeUsers = allFreeUsers.size;
      stats.totalPaidUsers = allPaidUsers.size;
      stats.breakdown.free.totalUsers = allFreeUsers.size;
      stats.breakdown.paid.totalUsers = allPaidUsers.size;
      
      // Convert monthly free/paid users to counts
      for (const [month, userSet] of Object.entries(freeUsersByMonth)) {
        stats.breakdown.free.monthlyActiveUsers[month] = userSet.size;
        if (month === stats.currentMonth) {
          stats.breakdown.free.currentMonthUsers = userSet.size;
        }
      }
      
      for (const [month, userSet] of Object.entries(paidUsersByMonth)) {
        stats.breakdown.paid.monthlyActiveUsers[month] = userSet.size;
        if (month === stats.currentMonth) {
          stats.breakdown.paid.currentMonthUsers = userSet.size;
        }
      }
      
      // Get auth keys for additional company info
      const authKeys = await redis.smembers('auth_keys') || [];
      for (const key of authKeys) {
        const keyData = await redis.hgetall(`auth_key:${key}`);
        if (keyData && keyData.company) {
          if (!stats.companies[keyData.company]) {
            stats.companies[keyData.company] = {
              monthlyActiveUsers: {},
              totalUniqueUsers: 0,
              currentMonthUsers: 0
            };
          }
          stats.companies[keyData.company].authKey = key;
          stats.companies[keyData.company].isActive = keyData.isActive;
        }
      }
      
      return res.status(200).json({ 
        success: true,
        stats: stats,
        message: '회사별 월별 사용자 통계'
      });
    } else {
      // No Redis available
      return res.status(200).json({ 
        success: true,
        stats: stats,
        message: 'Redis not configured - stats not available'
      });
    }
  } catch (error) {
    console.error('Usage stats error:', error);
    return res.status(500).json({ 
      error: '통계 조회 중 오류가 발생했습니다.',
      details: error.message
    });
  }
}