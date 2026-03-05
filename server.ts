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
  'https://www.googleapis.com/auth/calendar.events',
  'https://www.googleapis.com/auth/calendar',
  'https://www.googleapis.com/auth/spreadsheets'
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

app.get('/api/next-wp', async (req, res) => {
  const tokens = req.cookies.google_tokens;
  if (!tokens) return res.status(401).json({ error: 'Not authenticated' });

  try {
    const auth = new google.auth.OAuth2();
    auth.setCredentials(JSON.parse(tokens));
    const sheets = google.sheets({ version: 'v4', auth });
    const spreadsheetId = process.env.GOOGLE_SHEET_ID;

    if (!spreadsheetId) return res.json({ nextWp: '001-69' });

    const sheetData = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: 'Sheet1!C:C',
    });
    
    const rows = sheetData.data.values || [];
    const lastWp = rows.length > 0 ? rows[rows.length - 1][0] : null;
    
    const now = new Date();
    const currentYearBE = (now.getFullYear() + 543) % 100;
    const yearStr = currentYearBE.toString().padStart(2, '0');

    if (lastWp && typeof lastWp === 'string' && lastWp.includes('-')) {
      const [numPart, yearPart] = lastWp.split('-');
      const lastNum = parseInt(numPart, 10);
      const lastYear = parseInt(yearPart, 10);

      if (lastYear === currentYearBE) {
        return res.json({ nextWp: `${(lastNum + 1).toString().padStart(3, '0')}-${yearStr}` });
      }
    }
    res.json({ nextWp: `001-${yearStr}` });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch next WP' });
  }
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
    const sheets = google.sheets({ version: 'v4', auth });

    // 1. Upload to Google Drive
    const newFileName = `WP ผจฟ.1 No.${wpNumber} กสฟ.(ก3) เข้า ${stationName} (${date})`;
    
    const driveResponse = await drive.files.create({
      requestBody: {
        name: newFileName,
        parents: ['1aDRqNGcl934p1wzyuFsxd3bdX6PF6CD2'], 
      },
      media: {
        mimeType: file.mimetype,
        body: fs.createReadStream(file.path),
      },
      fields: 'id, name, webViewLink',
    });

    // 2. Find or Create "AI_WP" Calendar
    let calendarId = 'primary';
    try {
      const calendarList = await calendar.calendarList.list();
      const aiWpCalendar = calendarList.data.items?.find(c => c.summary === 'AI_WP');
      
      if (aiWpCalendar) {
        calendarId = aiWpCalendar.id!;
      } else {
        const newCalendar = await calendar.calendars.insert({
          requestBody: {
            summary: 'AI_WP',
            timeZone: 'Asia/Bangkok'
          }
        });
        calendarId = newCalendar.data.id!;
      }
    } catch (err) {
      console.error('Error finding/creating AI_WP calendar:', err);
      // Fallback to primary if error
    }

    // 3. Create Calendar Event
    const isoDate = req.body.isoDate || new Date().toISOString().split('T')[0];
    const startTime = req.body.startTime || `${isoDate}T08:00:00`;
    const endTime = req.body.endTime || `${isoDate}T17:00:00`;

    const event = {
      summary: calendarTitle,
      description: `Automated entry for ${newFileName}\nWork: ${req.body.workDescription}\nUnit: ${req.body.requestingUnit}`,
      start: {
        dateTime: startTime.includes('T') ? startTime : `${startTime}:00`,
        timeZone: 'Asia/Bangkok',
      },
      end: {
        dateTime: endTime.includes('T') ? endTime : `${endTime}:00`,
        timeZone: 'Asia/Bangkok',
      },
    };

    const calendarResponse = await calendar.events.insert({
      calendarId: calendarId,
      requestBody: event,
    });

    // 4. Write to Google Sheet
    let sheetLink = '';
    const spreadsheetId = process.env.GOOGLE_SHEET_ID;
    let finalWpNumber = wpNumber;

    if (spreadsheetId) {
      try {
        // Get the last WP number to calculate the next one
        const sheetData = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: 'Sheet1!C:C',
        });
        
        const rows = sheetData.data.values || [];
        const lastWp = rows.length > 0 ? rows[rows.length - 1][0] : null;
        
        const now = new Date();
        const currentYearBE = (now.getFullYear() + 543) % 100;
        const yearStr = currentYearBE.toString().padStart(2, '0');

        if (lastWp && typeof lastWp === 'string' && lastWp.includes('-')) {
          const [numPart, yearPart] = lastWp.split('-');
          const lastNum = parseInt(numPart, 10);
          const lastYear = parseInt(yearPart, 10);

          if (lastYear === currentYearBE) {
            finalWpNumber = `${(lastNum + 1).toString().padStart(3, '0')}-${yearStr}`;
          } else {
            finalWpNumber = `001-${yearStr}`;
          }
        } else {
          // Fallback if sheet is empty or format is wrong
          finalWpNumber = `001-${yearStr}`;
        }

        const timestamp = now.toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });
        const values = [
          [
            timestamp, // Column A: Upload Date/Time
            req.body.requestingUnit, // Column B: Requesting Unit
            finalWpNumber, // Column C: Run Number (Calculated)
            req.body.isStaffed, // Column D: Staffed/Unstaffed
            req.body.workDescription, // Column E: Work Description
            req.body.startTime, // Column F: Start Date/Time
            req.body.endTime, // Column G: End Date/Time
            req.body.department // Column H: Department
          ]
        ];

        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: 'Sheet1!A:H',
          valueInputOption: 'USER_ENTERED',
          requestBody: { values }
        });
        sheetLink = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
      } catch (err) {
        console.error('Error writing to Google Sheet:', err);
      }
    }

    // Update file name and calendar with the final calculated WP number
    const finalFileName = `WP ผจฟ.1 No.${finalWpNumber} กสฟ.(ก3) เข้า ${stationName} (${date})`;
    
    // Update Drive file name
    await drive.files.update({
      fileId: driveResponse.data.id!,
      requestBody: { name: finalFileName }
    });

    // Cleanup local file
    fs.unlinkSync(file.path);

    res.json({
      success: true,
      driveFile: { ...driveResponse.data, name: finalFileName },
      calendarEvent: calendarResponse.data,
      sheetLink: sheetLink,
      finalWpNumber: finalWpNumber
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
