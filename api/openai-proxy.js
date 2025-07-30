// Vercel Serverless Function
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
// 환경 변수가 없을 때 기본값에 GitHub Pages URL 포함
const ALLOWED_ORIGINS = process.env.ALLOWED_ORIGINS?.split(',').map(origin => origin.trim()) || [
  'https://localhost:3000',
  'https://excel.office.com',
  'https://latemonk.github.io'
];

// Import Upstash Redis
import { Redis } from '@upstash/redis';

let redis = null;
try {
  if (process.env.UPSTASH_REDIS_REST_URL && process.env.UPSTASH_REDIS_REST_TOKEN) {
    redis = new Redis({
      url: process.env.UPSTASH_REDIS_REST_URL,
      token: process.env.UPSTASH_REDIS_REST_TOKEN,
    });
    console.log('Redis connected successfully');
  } else {
    console.log('Redis environment variables not found, using fallback');
  }
} catch (error) {
  console.log('Redis initialization failed, using environment variables:', error);
}

// Valid auth keys - fallback to environment variable if KV not available
const VALID_AUTH_KEYS = process.env.VALID_AUTH_KEYS?.split(',').map(key => key.trim()) || [];

// Function to validate auth key and log validation
async function isValidAuthKey(authKey, authEmail, req) {
  console.log('isValidAuthKey called with:', { authKey, authEmail, hasRedis: !!redis });
  
  if (!authKey) return { valid: false, company: null };
  
  let company = null;
  let valid = false;
  
  // Try Redis first
  if (redis) {
    try {
      // First check if key exists in set
      const keyExists = await redis.sismember('auth_keys', authKey);
      console.log('Key exists in auth_keys set:', keyExists);
      
      const keyData = await redis.hgetall(`auth_key:${authKey}`);
      console.log('Redis lookup result:', JSON.stringify(keyData));
      console.log('isActive value:', keyData?.isActive, 'type:', typeof keyData?.isActive);
      
      if (keyData && keyData.isActive === 'true') {  // Redis returns strings
        valid = true;
        company = keyData.company || 'Unknown';
        
        // 사용 횟수 증가
        await redis.hincrby(`auth_key:${authKey}`, 'usageCount', 1);
        await redis.hset(`auth_key:${authKey}`, { lastUsed: new Date().toISOString() });
        
        // 로그 저장
        const logEntry = {
          authKey,
          email: authEmail || 'Not provided',
          company,
          timestamp: new Date().toISOString(),
          koreanTime: new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' }),
          ip: req.headers['x-forwarded-for'] || req.headers['x-real-ip'] || req.connection?.remoteAddress || 'Unknown',
          userAgent: req.headers['user-agent'] || 'Unknown',
          os: extractOS(req.headers['user-agent']),
          browser: extractBrowser(req.headers['user-agent']),
          origin: req.headers.origin || 'Unknown',
          model: req.body?.model || 'Unknown'
        };
        
        // Store log in Redis
        const logKey = `log:${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
        await redis.hset(logKey, logEntry);
        await redis.sadd('validation_logs', logKey);
        await redis.expire(logKey, 30 * 24 * 60 * 60); // Keep logs for 30 days
      }
    } catch (error) {
      console.error('Redis lookup error:', error);
    }
  }
  
  // Fallback to environment variable
  if (!valid && VALID_AUTH_KEYS.length > 0) {
    valid = VALID_AUTH_KEYS.includes(authKey);
    company = 'Demo/Test';
    console.log('Environment variable check:', { valid, VALID_AUTH_KEYS });
  }
  
  console.log('Final validation result:', { valid, company });
  return { valid, company };
}

// Helper functions to extract OS and Browser info
function extractOS(userAgent) {
  if (!userAgent) return 'Unknown';
  
  const osPatterns = [
    { pattern: /Windows NT 10.0/, name: 'Windows 10' },
    { pattern: /Windows NT 6.3/, name: 'Windows 8.1' },
    { pattern: /Windows NT 6.2/, name: 'Windows 8' },
    { pattern: /Windows NT 6.1/, name: 'Windows 7' },
    { pattern: /Mac OS X/, name: 'macOS' },
    { pattern: /Linux/, name: 'Linux' },
    { pattern: /iPhone/, name: 'iOS' },
    { pattern: /iPad/, name: 'iPadOS' },
    { pattern: /Android/, name: 'Android' }
  ];
  
  for (const { pattern, name } of osPatterns) {
    if (pattern.test(userAgent)) return name;
  }
  
  return 'Unknown';
}

function extractBrowser(userAgent) {
  if (!userAgent) return 'Unknown';
  
  const browserPatterns = [
    { pattern: /Edg/, name: 'Edge' },
    { pattern: /Chrome/, name: 'Chrome' },
    { pattern: /Safari/, name: 'Safari' },
    { pattern: /Firefox/, name: 'Firefox' },
    { pattern: /Opera|OPR/, name: 'Opera' }
  ];
  
  for (const { pattern, name } of browserPatterns) {
    if (pattern.test(userAgent)) return name;
  }
  
  return 'Unknown';
}

// CORS validation function
function isOriginAllowed(origin) {
  if (!origin) return false;
  
  // Check exact matches
  if (ALLOWED_ORIGINS.includes(origin)) return true;
  
  // Check wildcard
  if (ALLOWED_ORIGINS.includes('*')) return true;
  
  // Check patterns (e.g., for Office domains)
  const officePatterns = [
    /^https:\/\/.*\.office\.com$/,
    /^https:\/\/.*\.office365\.com$/,
    /^https:\/\/.*\.microsoft\.com$/,
    /^https:\/\/.*\.officeapps\.live\.com$/,
    /^https:\/\/.*\.sharepoint\.com$/,
    /^https:\/\/localhost:\d+$/
  ];
  
  return officePatterns.some(pattern => pattern.test(origin));
}

export default async function handler(req, res) {
  // CORS headers
  const origin = req.headers.origin;
  
  // Handle OPTIONS request first (preflight)
  if (req.method === 'OPTIONS') {
    if (isOriginAllowed(origin)) {
      res.setHeader('Access-Control-Allow-Origin', origin);
      res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
      res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
      res.setHeader('Access-Control-Max-Age', '86400'); // 24 hours
    }
    res.status(200).end();
    return;
  }
  
  // Handle GET request for health check (no origin check needed)
  if (req.method === 'GET') {
    res.status(200).json({ 
      status: 'ok',
      message: 'Excel Addon Proxy API is running',
      apiKeyConfigured: !!OPENAI_API_KEY,
      timestamp: new Date().toISOString(),
      version: '2.0', // Added to verify deployment
      authConfig: {
        redisConfigured: !!redis,
        envKeysCount: VALID_AUTH_KEYS.length,
        hasHardcodedKeys: false // This should be false after our fix
      }
    });
    return;
  }

  // For other requests, check origin
  console.log('Request method:', req.method);
  console.log('Request origin:', origin);
  console.log('User-Agent:', req.headers['user-agent']);
  
  if (isOriginAllowed(origin)) {
    res.setHeader('Access-Control-Allow-Origin', origin);
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  } else {
    console.log('Origin not allowed:', origin);
    console.log('Allowed origins:', ALLOWED_ORIGINS);
    res.status(403).json({ 
      error: 'Origin not allowed',
      origin: origin,
      allowedOrigins: ALLOWED_ORIGINS,
      headers: req.headers
    });
    return;
  }

  // Only allow POST for actual API calls
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  // Check if API key is configured
  if (!OPENAI_API_KEY) {
    console.error('OPENAI_API_KEY is not configured');
    res.status(500).json({
      success: false,
      error: 'API 키가 설정되지 않았습니다. 서버 설정을 확인해주세요.'
    });
    return;
  }

  try {
    const { command, sheetContext, model, authKey, authEmail } = req.body;
    
    console.log('Request received:', { 
      command, 
      hasSheetContext: !!sheetContext, 
      model, 
      hasAuthKey: !!authKey,
      authKeyLength: authKey?.length,
      authEmail 
    });

    if (!command || !sheetContext) {
      res.status(400).json({
        success: false,
        error: '잘못된 요청입니다.'
      });
      return;
    }
    
    // Validate auth key for premium model
    const selectedModel = model || 'gpt-4.1-mini-2025-04-14';
    console.log('Model validation check:', { selectedModel, requiresAuth: selectedModel === 'gpt-4.1-2025-04-14' });
    
    if (selectedModel === 'gpt-4.1-2025-04-14') {
      const validation = await isValidAuthKey(authKey, authEmail, req);
      if (!validation.valid) {
        console.log('Auth validation failed, returning 403');
        res.status(403).json({
          success: false,
          error: '프리미엄 모델을 사용하려면 유효한 인증키가 필요합니다.',
          debug: {
            authKeyProvided: !!authKey,
            authKeyLength: authKey?.length,
            redisAvailable: !!redis,
            envKeysCount: VALID_AUTH_KEYS.length,
            timestamp: new Date().toISOString()
          }
        });
        return;
      }
      console.log('Auth validation passed');
    }

    // Special handling for batch translation
    if (sheetContext.operation === 'translate_batch') {
      // Only validate and log for premium model
      if (selectedModel === 'gpt-4.1-2025-04-14' && authKey) {
        const validation = await isValidAuthKey(authKey, authEmail, req);
        if (!validation.valid) {
          res.status(403).json({
            success: false,
            error: '프리미엄 모델을 사용하려면 유효한 인증키가 필요합니다.'
          });
          return;
        }
      }
      const result = await translateBatch(sheetContext, selectedModel);
      res.status(200).json(result);
      return;
    }

    // Regular command processing
    const systemPrompt = `You are an Excel assistant that interprets natural language commands and returns JSON instructions for Excel operations.
    
Available operations:
1. merge: Merge cells
2. sum: Sum values in a range or column
3. average: Calculate average
4. count: Count cells (can count all, numbers only, or based on conditions)
5. format: Format cells (bold, italic, font color, background color, etc.)
6. sort: Sort data
7. filter: Filter data
8. insert: Insert rows/columns
9. delete: Delete rows/columns
10. formula: Add custom formula
11. chart: Create chart
12. conditional_format: Add conditional formatting
13. translate: Translate cell contents to another language
14. compress: Remove empty rows in a specific column range
15. remove_border: Remove cell borders
16. border_format: Format cell borders (color, style, etc.)

For count operation, parameters should include:
- "sourceRange": range to count from
- "targetCell": where to put the result (optional)
- "countType": "count", "counta", or "countif"
- "condition": for countif
- "operator": "contains", "equals", ">", "<", etc.

For sum operation:
- If user mentions a column by header name (e.g., "totalToken 열의 합", "totalToken 합산"), return: { "sumType": "column", "columnName": "totalToken" }
- If user mentions a column by letter (e.g., "D열 합계", "D column sum"), return: { "sumType": "column", "columnName": "D" }
- For specific range sum, use: { "sourceRange": "A2:A10" }
- For adding sum below selection, use: { "addNewRow": true }

For average operation:
- If user mentions column average (e.g., "C열 평균"), return: { "averageType": "column", "column": "C" }
- If user mentions row average (e.g., "3행 평균"), return: { "averageType": "row", "row": 3 }
- If user mentions range average (e.g., "C1-C100 평균", "C1:C100 평균"), return: { "sourceRange": "C1:C100" }
- If user mentions column by header name (e.g., "총액 평균"), return: { "averageType": "column", "columnName": "총액" }
- Default behavior without specific range uses selected cells

For format operation:
- If user mentions number format (e.g., "숫자 형식", "숫자로"), return: { "numberFormat": "number" }
- If user mentions currency/won format (e.g., "원화 형식", "통화 형식"), return: { "numberFormat": "currency" }
- For specific cells like "E101", use: { "range": "E101" }
- If user mentions text color (e.g., "파란색으로", "빨간색 글자"), use: { "fontColor": "#0000FF" } (not backgroundColor)
- If user mentions background/cell color (e.g., "배경색", "셀 색상"), use: { "backgroundColor": "#color" }
- Other format options: bold (굵게), italic (기울임), fontSize (글자 크기)
- Common colors: 파란색=#0000FF, 빨간색=#FF0000, 초록색=#00FF00, 노란색=#FFFF00, 검정색=#000000

For sort operation:
- If user mentions column by letter (e.g., "B열"), extract column number (B=2, C=3, etc.)
- If user mentions "내림차순" or "큰 순서대로", use: { "ascending": false }
- If user mentions "오름차순" or "작은 순서대로", use: { "ascending": true }
- Default is ascending if not specified
- Parameters: { "column": 2, "ascending": false }

For conditional_format operation:
- IMPORTANT: Do NOT include "range" parameter unless user specifically mentions a range
- If user just says "값이 X보다 큰 셀" without specifying range, do NOT add range parameter
- If user mentions numeric comparison (e.g., "100보다 큰", "50 미만"), it will apply to all cells in the sheet
- Conditions: "greater_than" (크다, 초과), "less_than" (작다, 미만), "equal_to" (같다)
- Colors: use hex values like "#00FF00" for green (녹색), "#FF0000" for red (빨간색)
- Example (NO range): { "condition": "greater_than", "value": 100, "backgroundColor": "#00FF00" }
- Only add range if user says something like "A열에서", "선택한 범위에서", etc.

For chart operation:
- IMPORTANT: If user just says "차트" or "그래프" without specifying type, ALWAYS use: { "chartType": "bar" }
- If user mentions specific chart type (e.g., "막대 차트", "선 그래프"), use the specified type
- Chart types: "bar" (막대, DEFAULT), "line" (선), "pie" (원), "scatter" (분산형)
- For specific range like "A1:B10", use: { "range": "A1:B10" }
- If active range shows multiple non-contiguous cells (e.g., "B2,C10,D12"), the chart will consolidate the data automatically
- For individual cells, the system will create a temporary consolidated range
- Example: { "chartType": "bar", "title": "데이터 차트" }
- ALWAYS include chartType parameter, default to "bar" if not specified

For translate operation:
- If user mentions column by letter (e.g., "C열을 영어로 번역"), use: { "sourceRange": "C:C", "targetLanguage": "영어" }
- If user mentions specific range (e.g., "B2-B40", "B2:B40"), use: { "sourceRange": "B2:B40", "targetLanguage": "일본어" }
- If user specifies target column (e.g., "E열에 추가", "E열에 넣어"), use: { "targetRange": "E:E" }
- IMPORTANT: Target column must be extracted from phrases like "E열에", "E column", "to column E"
- Languages: 영어 (English), 한국어 (Korean), 일본어 (Japanese), 중국어 (Chinese), etc.
- Example: { "sourceRange": "B2:B40", "targetRange": "E:E", "targetLanguage": "일본어" }

For compress operation:
- If user mentions removing empty rows in a range (e.g., "D2:D170 사이의 빈 행 제거"), use: { "range": "D2:D170" }
- This removes entire rows where the specified column cells are empty
- Example: { "range": "D2:D170" }

For remove_border operation:
- If user mentions removing border (e.g., "테두리 없애", "border 제거"), use: { "borderType": "all" }
- Border types: "all" (모든 테두리), "right" (오른쪽), "left" (왼쪽), "top" (위), "bottom" (아래)
- If user specifies a specific column (e.g., "C열의 오른쪽 테두리"), use: { "range": "C:C", "borderType": "right" }
- If user says "모든 셀", "전체 시트", "시트 전체", "모든 열", "모든 행", use: { "range": "all", "borderType": "all" }
- If user says "선택한" or "지정한" (e.g., "선택한 셀", "지정한 열", "선택한 범위", "지정한 행"), don't include range parameter (uses selected range)
- Example: { "range": "C:C", "borderType": "right" } or { "range": "all", "borderType": "all" }

For border_format operation:
- If user mentions changing border color (e.g., "테두리 빨간색으로", "border 파란색"), use border_format operation
- Border types: "all" (모든 테두리), "right" (오른쪽), "left" (왼쪽), "top" (위), "bottom" (아래), "inside" (내부)
- Colors: "빨간색"="#FF0000", "파란색"="#0000FF", "검정색"="#000000", "초록색"="#00FF00", "노란색"="#FFFF00"
- Border styles: "continuous" (실선, default), "dash" (점선), "dashDot" (일점쇄선), "double" (이중선)
- Example: { "borderType": "all", "color": "#FF0000", "style": "continuous" }
- If user says "선택한" or row/range is mentioned, don't include range parameter (uses selected range)

Current sheet context:
- Active range: ${sheetContext.activeRange?.address}
- Sheet dimensions: ${sheetContext.lastRow} rows x ${sheetContext.lastColumn} columns
- Headers: ${sheetContext.headers?.map(h => `Column ${h.columnLetter}: "${h.label}"`).join(', ')}

IMPORTANT: If user requests multiple operations in one command (e.g., "format column A as number and column B as currency"), 
return an array of operations in this format:
{
  "operations": [
    {
      "operation": "format",
      "parameters": { "columnName": "totalTokens", "numberFormat": "number" }
    },
    {
      "operation": "format", 
      "parameters": { "columnName": "totalCharge", "numberFormat": "currency" }
    }
  ]
}

For single operations, return:
{
  "operation": "operation_name",
  "parameters": {
    // operation-specific parameters
  }
}`;

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${OPENAI_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: selectedModel,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: `User command: ${command}` }
        ],
        temperature: 0.3,
        max_tokens: 500
      })
    });

    if (!response.ok) {
      const error = await response.json();
      if (response.status === 429) {
        res.status(200).json({
          success: false,
          error: 'API 요청 한도를 초과했습니다. 잠시 후 다시 시도해주세요.'
        });
        return;
      }
      res.status(200).json({
        success: false,
        error: `API 오류: ${error.error?.message || '알 수 없는 오류'}`
      });
      return;
    }

    const data = await response.json();
    
    if (data.choices && data.choices[0]) {
      const content = data.choices[0].message.content;
      try {
        const parsedCommand = JSON.parse(content);
        res.status(200).json({
          success: true,
          data: parsedCommand
        });
      } catch (parseError) {
        res.status(200).json({
          success: false,
          error: 'AI 응답을 해석할 수 없습니다.'
        });
      }
    } else {
      res.status(200).json({
        success: false,
        error: 'OpenAI API 응답을 파싱할 수 없습니다.'
      });
    }
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({
      success: false,
      error: '서버 오류가 발생했습니다.'
    });
  }
}

// Handle batch translation
async function translateBatch(context, model = 'gpt-4.1-mini-2025-04-14') {
  const { texts, targetLanguage, sourceLanguage } = context;
  
  const numberedTexts = texts.map((text, index) => `[${index + 1}] ${text}`);
  
  // Map target languages to their specific language codes for clarity
  const languageMap = {
    '영어': 'English',
    '일본어': 'Japanese', 
    '중국어': 'Chinese',
    '한국어': 'Korean',
    '스페인어': 'Spanish',
    '프랑스어': 'French',
    '독일어': 'German'
  };
  
  const targetLangCode = languageMap[targetLanguage] || targetLanguage;
  const sourceLangCode = sourceLanguage ? (languageMap[sourceLanguage] || sourceLanguage) : 'auto-detect';
  
  const systemPrompt = `You are a professional translator for spreadsheet data. Your task is to translate text ONLY to ${targetLangCode}.

CRITICAL RULES:
1. ALWAYS translate ALL items to ${targetLangCode} ONLY - never use any other language
2. Each numbered item MUST be translated separately to ${targetLangCode}
3. Return translations in EXACT format: [1] ${targetLangCode}_translation\n[2] ${targetLangCode}_translation\n...
4. If an item is empty or untranslatable, return [N] [EMPTY] for that number
5. NEVER translate to Korean unless ${targetLangCode} is specifically Korean
6. NEVER keep the original language - always translate to ${targetLangCode}
7. You MUST return EXACTLY ${texts.length} numbered translations - no more, no less
8. NEVER skip numbers - if you receive [1] through [20], you MUST return [1] through [20]
9. Each batch is independent - maintain ${targetLangCode} consistency throughout ALL items
10. IMPORTANT: If source is Korean and target is Japanese, you MUST translate Korean text to Japanese
11. DO NOT return the original Korean text - it must be translated to ${targetLangCode}`;

  const userPrompt = sourceLanguage 
    ? `Translate ALL of these ${texts.length} items from ${sourceLangCode} to ${targetLangCode} (IMPORTANT: Every single item must be in ${targetLangCode}, not ${sourceLangCode} or any other language):\n\n${numberedTexts.join('\n')}`
    : `Translate ALL of these ${texts.length} items to ${targetLangCode} (IMPORTANT: Every single item must be in ${targetLangCode}, not the original language or Korean):\n\n${numberedTexts.join('\n')}`;

  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${OPENAI_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: model,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userPrompt }
        ],
        temperature: 0.3,
        max_tokens: 2000
      })
    });

    if (!response.ok) {
      const error = await response.json();
      console.error('OpenAI translation error:', error);
      return {
        success: false,
        error: `Translation API error: ${error.error?.message || 'Unknown error'}`
      };
    }
    
    const data = await response.json();
    
    if (data.choices && data.choices[0]) {
      const responseText = data.choices[0].message.content.trim();
      const translations = [];
      const lines = responseText.split('\n');
      
      const translationMap = {};
      for (const line of lines) {
        const match = line.match(/^\[(\d+)\]\s*(.*)$/);
        if (match) {
          const num = parseInt(match[1]);
          const translation = match[2].trim();
          translationMap[num] = translation === '[EMPTY]' ? '' : translation;
        }
      }
      
      // Ensure we have exactly the same number of translations as input texts
      for (let i = 1; i <= texts.length; i++) {
        if (translationMap.hasOwnProperty(i)) {
          translations.push(translationMap[i]);
        } else {
          // If translation is missing, mark as empty to maintain row alignment
          console.log(`Warning: Missing translation for item ${i}`);
          translations.push('');
        }
      }
      
      return {
        success: true,
        data: {
          operation: 'translate_batch_result',
          translations: translations
        }
      };
    } else {
      console.error('Invalid OpenAI response:', data);
      return {
        success: false,
        error: 'Invalid response from translation API'
      };
    }
  } catch (error) {
    console.error('Translation error:', error);
    return {
      success: false,
      error: '번역 중 오류가 발생했습니다.'
    };
  }
}