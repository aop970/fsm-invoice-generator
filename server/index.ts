/**
 * index.ts — Express server for FSM Invoice Generator
 * POST /generate → fills template → returns .xlsx buffer
 */

import express from 'express';
import type { Request, Response } from 'express';
import { fillTemplate } from './filler';
import type { GeneratePayload } from './filler';

const app = express();
const PORT = 3001;

// Parse JSON bodies (invoices can be large — 50mb limit covers thousands of rows)
app.use(express.json({ limit: '50mb' }));

app.post('/generate', async (req: Request, res: Response): Promise<void> => {
  try {
    const payload = req.body as GeneratePayload;

    if (!payload || !payload.fsmI || !payload.fsmII) {
      res.status(400).json({ error: 'Invalid payload: fsmI and fsmII are required' });
      return;
    }

    const buf = await fillTemplate(payload);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${payload.invoiceName}.xlsx"`);
    res.send(buf);
  } catch (err) {
    console.error('[/generate] Error:', err);
    const msg = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: msg });
  }
});

app.listen(PORT, () => {
  console.log(`[FSM Server] Listening on http://localhost:${PORT}`);
});
