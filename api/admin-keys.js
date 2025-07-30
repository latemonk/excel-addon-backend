// Vercel Serverless Function for Auth Key Management
// This is an example - you'd need to add authentication to protect this endpoint

export default async function handler(req, res) {
  // IMPORTANT: Add authentication here in production!
  const adminPassword = req.headers['x-admin-password'];
  if (adminPassword !== process.env.ADMIN_PASSWORD) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  if (req.method === 'GET') {
    // Get all keys from environment variable
    const keys = process.env.VALID_AUTH_KEYS?.split(',').map(key => key.trim()) || [];
    return res.status(200).json({ keys });
  }

  if (req.method === 'POST') {
    const { action, key } = req.body;
    
    if (action === 'generate') {
      // Generate new key
      const newKey = 'KEY' + Math.random().toString(36).substr(2, 9).toUpperCase();
      
      // Note: In Vercel, you can't modify env vars at runtime
      // You'd need to use Vercel KV or a database for dynamic key management
      
      return res.status(200).json({ 
        key: newKey,
        message: 'Key generated. Add to VALID_AUTH_KEYS in Vercel dashboard.'
      });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
}