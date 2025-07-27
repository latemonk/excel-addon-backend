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
    
    // OpenAI API 호출
    const systemPrompt = `You are an Excel automation assistant. Analyze the command and return JSON with the operation to perform.
Available operations: merge, sum, average, count, format, sort, filter, insert, delete, formula, chart, conditional_format, translate

Current context:
- Active range: ${sheetContext.activeRange?.address}
- Sheet: ${sheetContext.sheetName}
- Headers: ${sheetContext.headers?.map(h => `Column ${h.columnLetter}: "${h.label}"`).join(', ') || 'No headers'}

For sum operation:
- If user mentions a column by header name (e.g., "totalToken 열의 합", "totalToken 합산"), return: { "sumType": "column", "columnName": "totalToken" }
- For specific range sum, use: { "sourceRange": "A2:A10" }
- For adding sum below selection, use: { "addNewRow": true }

For format operation:
- If user mentions number format (e.g., "숫자 형식", "숫자로"), return: { "numberFormat": "number" }
- If user mentions currency/won format (e.g., "원화 형식", "통화 형식"), return: { "numberFormat": "currency" }
- For specific cells like "E101", use: { "range": "E101" }
- Other format options: bold, italic, fontSize, fontColor, backgroundColor

Return JSON in format:
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