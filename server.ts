import express from 'express';
import path from 'path';
import { createServer as createViteServer } from 'vite';
import { GoogleGenAI } from '@google/genai';

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json());

  // API Routes
  app.post('/api/summarize', async (req, res) => {
    try {
      const { transcript } = req.body;
      if (!transcript) {
        return res.status(400).json({ error: 'Transcript is required' });
      }

      if (!process.env.GEMINI_API_KEY) {
        return res.status(500).json({ error: 'GEMINI_API_KEY is not configured.' });
      }

      const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: `You are an AI meeting assistant. Please provide a well-structured summary and key action items for the following meeting transcript. Use Traditional Chinese (zh-TW):\n\n${transcript}`,
      });

      res.json({ summary: response.text });
    } catch (err: any) {
      console.error('Error generating summary:', err);
      res.status(500).json({ error: err.message || 'Failed to generate summary' });
    }
  });

  app.post('/api/send-email', async (req, res) => {
    try {
      const { to, subject, text } = req.body;
      if (!to || !text) {
        return res.status(400).json({ error: '缺少必填欄位：收件者或內容' });
      }
      
      // Simulate network delay
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      // We can simulate an error if the email contains "error"
      if (to.includes('error')) {
         return res.status(400).json({ error: '無法寄送：無效的收件人地址。此為模擬錯誤。' });
      }

      console.log(`Sending simulated email to ${to}...`);
      
      res.json({ success: true, message: '信件已成功寄出！ (模擬)' });
    } catch (err: any) {
      console.error('Error sending email:', err);
      res.status(500).json({ error: err.message || '系統錯誤，無法寄送郵件' });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
