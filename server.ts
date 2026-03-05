import express from 'express';
import { createServer as createViteServer } from 'vite';
import { google } from 'googleapis';
import multer from 'multer';
import cookieParser from 'cookie-parser';
import path from 'path';
import fs from 'fs';

const app = express();
const PORT = 3000;

// Ensure uploads directory exists
if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}

const upload = multer({ dest: 'uploads/' });

app.use(express.json());
app.use(cookieParser());

// Google OAuth Configuration
const getRedirectUri = () => {
  const baseUrl = (process.env.APP_URL || '').replace(/\/$/, '');
  return `${baseUrl}/auth/callback`;
};

const oauth2Client = new google.auth.OAuth2(
  process.env.GOOGLE_CLIENT_ID,
  process.env.GOOGLE_CLIENT_SECRET,
  getRedirectUri()
);

const SCOPES = [
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/calendar.events'
];

// Auth Routes
app.get('/api/auth/url', (req, res) => {
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent'
  });
  res.json({ url });
});

app.get('/auth/callback', async (req, res) => {
  const { code } = req.query;
  try {
    const { tokens } = await oauth2Client.getToken(code as string);
    // Store tokens in a secure cookie
    res.cookie('google_tokens', JSON.stringify(tokens), {
      httpOnly: true,
      secure: true,
      sameSite: 'none',
      maxAge: 30 * 24 * 60 * 60 * 1000 // 30 days
    });

    res.send(`
      <html>
        <body>
          <script>
            if (window.opener) {
              window.opener.postMessage({ type: 'OAUTH_AUTH_SUCCESS' }, '*');
              window.close();
            } else {
              window.location.href = '/';
            }
          </script>
          <p>Authentication successful. This window should close automatically.</p>
        </body>
      </html>
    `);
  } catch (error) {
    console.error('Error getting tokens:', error);
    res.status(500).send('Authentication failed');
  }
});

app.get('/api/auth/status', (req, res) => {
  const tokens = req.cookies.google_tokens;
  res.json({ isAuthenticated: !!tokens });
});

app.post('/api/auth/logout', (req, res) => {
  res.clearCookie('google_tokens');
  res.json({ success: true });
});

// Drive & Calendar Processing
app.post('/api/process', upload.single('file'), async (req: any, res) => {
  const tokens = req.cookies.google_tokens;
  if (!tokens) return res.status(401).json({ error: 'Not authenticated' });

  const { wpNumber, stationName, date, calendarTitle } = req.body;
  const file = req.file;

  if (!file) return res.status(400).json({ error: 'No file uploaded' });

  try {
    const auth = new google.auth.OAuth2();
    auth.setCredentials(JSON.parse(tokens));

    const drive = google.drive({ version: 'v3', auth });
    const calendar = google.calendar({ version: 'v3', auth });

    // 1. Upload to Google Drive
    // Pattern: WP ผจฟ.1 No.013-69 กสฟ.(ก3) เข้า สถานีไฟฟ้ากระทุ่มแบน 6 (13 ก.พ. 69)
    const newFileName = `WP ผจฟ.1 No.${wpNumber} กสฟ.(ก3) เข้า ${stationName} (${date})`;
    
    const driveResponse = await drive.files.create({
      requestBody: {
        name: newFileName,
        parents: ['1aDRqNGcl934p1wzyuFsxd3bdX6PF6CD2'], // The specific folder ID
      },
      media: {
        mimeType: file.mimetype,
        body: fs.createReadStream(file.path),
      },
      fields: 'id, name, webViewLink',
    });

    // 2. Create Calendar Event
    // Pattern: สถานีไฟฟ้านครชัยศรี 1 (จัดพนักงาน) WP ผจฟ.1 No.100 บำรุงรักษาระบบ SCPS ประจำปี (ผปค.กสฟ.ก3)
    // We need to parse the date string. Gemini should provide a standard ISO date for the API.
    // For now, let's assume Gemini gives us a start and end date or we use the extracted date.
    
    // Attempt to parse the date from the user-friendly string (e.g., "13 ก.พ. 69")
    // However, it's safer if Gemini provides an ISO date in the background.
    const isoDate = req.body.isoDate || new Date().toISOString().split('T')[0];

    const event = {
      summary: calendarTitle,
      description: `Automated entry for ${newFileName}`,
      start: {
        date: isoDate,
      },
      end: {
        date: isoDate,
      },
    };

    const calendarResponse = await calendar.events.insert({
      calendarId: 'primary',
      requestBody: event,
    });

    // Cleanup local file
    fs.unlinkSync(file.path);

    res.json({
      success: true,
      driveFile: driveResponse.data,
      calendarEvent: calendarResponse.data,
    });
  } catch (error: any) {
    console.error('Processing error:', error);
    if (file) fs.unlinkSync(file.path);
    res.status(500).json({ error: error.message });
  }
});

async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static('dist'));
    app.get('*', (req, res) => {
      res.sendFile(path.resolve('dist/index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
