// Azure Functions 백엔드 (별도 배포 필요)
// 파일명: openai-proxy/index.js

const axios = require('axios');

// 환경 변수에서 API 키 읽기
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const ALLOWED_ORIGINS = process.env.ALLOWED_ORIGINS?.split(',') || ['https://localhost:3000'];

module.exports = async function (context, req) {
    // CORS 설정
    const origin = req.headers.origin;
    if (ALLOWED_ORIGINS.includes(origin)) {
        context.res = {
            headers: {
                'Access-Control-Allow-Origin': origin,
                'Access-Control-Allow-Methods': 'POST, OPTIONS',
                'Access-Control-Allow-Headers': 'Content-Type',
                'Access-Control-Max-Age': '86400'
            }
        };
    }

    // OPTIONS 요청 처리
    if (req.method === 'OPTIONS') {
        context.res.status = 200;
        return;
    }

    // POST 요청 처리
    if (req.method === 'POST') {
        try {
            const { command, sheetContext } = req.body;

            if (!command || !sheetContext) {
                context.res = {
                    status: 400,
                    body: {
                        success: false,
                        error: '잘못된 요청입니다.'
                    }
                };
                return;
            }

            // OpenAI API 호출
            const response = await callOpenAI(command, sheetContext);
            
            context.res = {
                status: 200,
                body: response,
                headers: {
                    'Content-Type': 'application/json',
                    ...context.res.headers
                }
            };
        } catch (error) {
            context.res = {
                status: 500,
                body: {
                    success: false,
                    error: error.message
                }
            };
        }
    } else {
        context.res = {
            status: 405,
            body: {
                success: false,
                error: 'Method not allowed'
            }
        };
    }
};

async function callOpenAI(command, sheetContext) {
    const systemPrompt = `You are an Excel assistant that interprets natural language commands and returns JSON instructions for Excel operations.
    
Available operations:
1. merge: Merge cells
2. sum: Sum values in a range
3. average: Calculate average
4. count: Count cells
5. format: Format cells
6. sort: Sort data
7. filter: Filter data
8. insert: Insert rows/columns
9. delete: Delete rows/columns
10. formula: Add custom formula
11. chart: Create chart
12. conditional_format: Add conditional formatting
13. translate: Translate cell contents

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

    try {
        const response = await axios.post('https://api.openai.com/v1/chat/completions', {
            model: 'gpt-4.1-2025-04-14',
            messages: [
                { role: 'system', content: systemPrompt },
                { role: 'user', content: `User command: ${command}` }
            ],
            temperature: 0.3,
            max_tokens: 500
        }, {
            headers: {
                'Authorization': `Bearer ${OPENAI_API_KEY}`,
                'Content-Type': 'application/json'
            }
        });

        if (response.data.choices && response.data.choices[0]) {
            const content = response.data.choices[0].message.content;
            try {
                const parsedCommand = JSON.parse(content);
                return {
                    success: true,
                    data: parsedCommand
                };
            } catch (parseError) {
                return {
                    success: false,
                    error: 'AI 응답을 해석할 수 없습니다.'
                };
            }
        }
    } catch (error) {
        if (error.response?.status === 429) {
            return {
                success: false,
                error: 'API 요청 한도를 초과했습니다. 잠시 후 다시 시도해주세요.'
            };
        }
        return {
            success: false,
            error: `API 오류: ${error.message}`
        };
    }
}