// Vercel Serverless Function
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
// 환경 변수가 없을 때 기본값에 GitHub Pages URL 포함
const ALLOWED_ORIGINS = process.env.ALLOWED_ORIGINS?.split(',').map(origin => origin.trim()) || [
  'https://localhost:3000',
  'https://excel.office.com',
  'https://latemonk.github.io'
];

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
      timestamp: new Date().toISOString()
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
    const { command, sheetContext } = req.body;

    if (!command || !sheetContext) {
      res.status(400).json({
        success: false,
        error: '잘못된 요청입니다.'
      });
      return;
    }

    // Special handling for batch translation
    if (sheetContext.operation === 'translate_batch') {
      const result = await translateBatch(sheetContext);
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

For count operation, parameters should include:
- "sourceRange": range to count from
- "targetCell": where to put the result (optional)
- "countType": "count", "counta", or "countif"
- "condition": for countif
- "operator": "contains", "equals", ">", "<", etc.

For sum operation:
- If user mentions a column by header name (e.g., "totalToken 열의 합", "totalToken 합산"), return: { "sumType": "column", "columnName": "totalToken" }
- For specific range sum, use: { "sourceRange": "A2:A10" }
- For adding sum below selection, use: { "addNewRow": true }

For format operation:
- If user mentions number format (e.g., "숫자 형식", "숫자로"), return: { "numberFormat": "number" }
- If user mentions currency/won format (e.g., "원화 형식", "통화 형식"), return: { "numberFormat": "currency" }
- For specific cells like "E101", use: { "range": "E101" }
- If user mentions text color (e.g., "파란색으로", "빨간색 글자"), use: { "fontColor": "#0000FF" } (not backgroundColor)
- If user mentions background/cell color (e.g., "배경색", "셀 색상"), use: { "backgroundColor": "#color" }
- Other format options: bold (굵게), italic (기울임), fontSize (글자 크기)
- Common colors: 파란색=#0000FF, 빨간색=#FF0000, 초록색=#00FF00, 노란색=#FFFF00, 검정색=#000000

For conditional_format operation:
- "condition": "greater_than", "less_than", "equal_to", "text_contains", "not_empty", "empty"
- "value": the value to compare
- "backgroundColor": hex color
- "fontColor": hex color
- "bold": true/false

Current sheet context:
- Active range: ${sheetContext.activeRange?.address}
- Sheet dimensions: ${sheetContext.lastRow} rows x ${sheetContext.lastColumn} columns
- Headers: ${sheetContext.headers?.map(h => `Column ${h.columnLetter}: "${h.label}"`).join(', ')}

Return JSON in this format:
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
        model: 'gpt-4.1-2025-04-14',
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
async function translateBatch(context) {
  const { texts, targetLanguage, sourceLanguage } = context;
  
  const numberedTexts = texts.map((text, index) => `[${index + 1}] ${text}`);
  
  const systemPrompt = `You are a professional translator for spreadsheet data. CRITICAL RULES:
1. Each numbered item MUST be translated separately
2. Return translations in EXACT same format: [1] translation1\\n[2] translation2\\n...
3. If an item is empty or untranslatable, return [N] [EMPTY] for that number
4. Maintain the exact count of items`;

  const userPrompt = sourceLanguage 
    ? `Translate these ${texts.length} items from ${sourceLanguage} to ${targetLanguage}:\\n\\n${numberedTexts.join('\\n')}`
    : `Translate these ${texts.length} items to ${targetLanguage}:\\n\\n${numberedTexts.join('\\n')}`;

  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${OPENAI_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: 'gpt-4.1-2025-04-14',
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userPrompt }
        ],
        temperature: 0.3,
        max_tokens: 2000
      })
    });

    const data = await response.json();
    
    if (data.choices && data.choices[0]) {
      const responseText = data.choices[0].message.content.trim();
      const translations = [];
      const lines = responseText.split('\\n');
      
      const translationMap = {};
      for (const line of lines) {
        const match = line.match(/^\\[(\\d+)\\]\\s*(.*)$/);
        if (match) {
          const num = parseInt(match[1]);
          const translation = match[2].trim();
          translationMap[num] = translation === '[EMPTY]' ? '' : translation;
        }
      }
      
      for (let i = 1; i <= texts.length; i++) {
        translations.push(translationMap[i] || '');
      }
      
      return {
        success: true,
        data: {
          operation: 'translate_batch_result',
          translations: translations
        }
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