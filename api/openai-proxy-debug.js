// Debug version - CORS를 일시적으로 완화한 버전
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;

export default async function handler(req, res) {
  const origin = req.headers.origin || req.headers.referer;
  
  // 디버깅을 위해 모든 요청 정보 로깅
  console.log('=== REQUEST DEBUG INFO ===');
  console.log('Method:', req.method);
  console.log('Origin:', origin);
  console.log('Headers:', JSON.stringify(req.headers, null, 2));
  console.log('=========================');
  
  // OPTIONS 요청 처리 (모든 origin 허용)
  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    res.setHeader('Access-Control-Max-Age', '86400');
    res.status(200).end();
    return;
  }
  
  // 모든 요청에 CORS 헤더 설정
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  
  // GET 요청 - 상태 확인
  if (req.method === 'GET') {
    res.status(200).json({ 
      status: 'ok',
      message: 'Debug API is running',
      apiKeyConfigured: !!OPENAI_API_KEY,
      requestOrigin: origin,
      timestamp: new Date().toISOString()
    });
    return;
  }
  
  // POST 요청만 처리
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }
  
  // API 키 확인
  if (!OPENAI_API_KEY) {
    res.status(500).json({
      success: false,
      error: 'API 키가 설정되지 않았습니다.'
    });
    return;
  }
  
  try {
    const { command, sheetContext } = req.body;
    
    console.log('Command:', command);
    console.log('Context:', JSON.stringify(sheetContext, null, 2));
    
    // Special handling for batch translation
    if (sheetContext.operation === 'translate_batch') {
      const result = await translateBatch(sheetContext);
      res.status(200).json(result);
      return;
    }
    
    // OpenAI API 호출
    const systemPrompt = `You are an Excel automation assistant. Analyze the command and return JSON with the operation to perform.
Available operations: merge, sum, average, count, format, sort, filter, insert, delete, formula, chart, conditional_format, translate, remove_border

Current context:
- Active range: ${sheetContext.activeRange?.address}
- Sheet: ${sheetContext.sheetName}
- Headers: ${sheetContext.headers?.map(h => `Column ${h.columnLetter}: "${h.label}"`).join(', ') || 'No headers'}

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
- If user mentions numeric comparison (e.g., "100보다 큰", "50 미만"), it will only apply to numeric cells
- Conditions: "greater_than" (크다, 초과), "less_than" (작다, 미만), "equal_to" (같다)
- Colors: use hex values like "#00FF00" for green (녹색), "#FF0000" for red (빨간색)
- Example: { "condition": "greater_than", "value": 100, "backgroundColor": "#00FF00" }

For chart operation:
- If user mentions chart/graph (e.g., "차트", "그래프", "막대 차트"), use: { "chartType": "bar" }
- Chart types: "bar" (막대), "line" (선), "pie" (원), "scatter" (분산형)
- For specific range like "A1:B10", use: { "range": "A1:B10" }
- Example: { "chartType": "bar", "range": "A1:B10", "title": "데이터 차트" }

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
  "parameters": {}
}`;
    
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${OPENAI_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: 'gpt-3.5-turbo',
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: command }
        ],
        temperature: 0.3,
        max_tokens: 500
      })
    });
    
    const data = await response.json();
    
    if (!response.ok) {
      console.error('OpenAI API error:', data);
      res.status(500).json({
        success: false,
        error: data.error?.message || 'OpenAI API 오류'
      });
      return;
    }
    
    const result = JSON.parse(data.choices[0].message.content);
    res.status(200).json({
      success: true,
      data: result
    });
    
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
}

// Handle batch translation
async function translateBatch(context) {
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
        model: 'gpt-3.5-turbo',
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
      console.log('Raw OpenAI response:', responseText);
      console.log('Response length:', responseText.length);
      
      const translations = [];
      const lines = responseText.split('\n');
      console.log('Number of lines:', lines.length);
      console.log('First 5 lines:', lines.slice(0, 5));
      
      const translationMap = {};
      for (const line of lines) {
        const match = line.match(/^\[(\d+)\]\s*(.*)$/);
        if (match) {
          const num = parseInt(match[1]);
          const translation = match[2].trim();
          translationMap[num] = translation === '[EMPTY]' ? '' : translation;
          console.log(`Parsed [${num}]: "${translation}"`);
        } else if (line.trim()) {
          console.log('Line did not match pattern:', line);
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