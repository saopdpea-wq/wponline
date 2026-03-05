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

const upload = multer({ 
  dest: 'uploads/',
  limits: { fileSize: 50 * 1024 * 1024 } // 50MB
});

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));
app.use(cookieParser());

// Google OAuth Configuration
const getRedirectUri = () => {
  const baseUrl = (process.env.APP_URL || '').replace(/\/$/, '');
  return `${baseUrl}/auth/callback`;
};

const getOAuth2Client = () => {
  const clientId = process.env.GOOGLE_CLIENT_ID || process.env.CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET || process.env.CLIENT_SECRET;
  
  if (!clientId || !clientSecret) {
    console.error('MISSING GOOGLE_CLIENT_ID or GOOGLE_CLIENT_SECRET in environment variables');
  }

  return new google.auth.OAuth2(
    clientId,
    clientSecret,
    getRedirectUri()
  );
};

const oauth2Client = getOAuth2Client();

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
    const auth = getOAuth2Client();
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

  const file = req.file;
  if (!file) return res.status(400).json({ error: 'No file uploaded' });

  let events = [];
  try {
    events = JSON.parse(req.body.events || '[]');
  } catch (e) {
    return res.status(400).json({ error: 'Invalid events data' });
  }

  if (events.length === 0) return res.status(400).json({ error: 'No events to process' });

  try {
    const auth = getOAuth2Client();
    auth.setCredentials(JSON.parse(tokens));

    const drive = google.drive({ version: 'v3', auth });
    const calendar = google.calendar({ version: 'v3', auth });
    const sheets = google.sheets({ version: 'v4', auth });

    // 1. Upload to Google Drive (Original File)
    const firstEvent = events[0];
    const driveFileName = `WP ผจฟ.1 (Multi-Entry) เข้า ${firstEvent.stationName} (${firstEvent.date})`;
    
    let driveResponse;
    try {
      driveResponse = await drive.files.create({
        requestBody: {
          name: driveFileName,
          parents: ['1aDRqNGcl934p1wzyuFsxd3bdX6PF6CD2'], 
        },
        media: {
          mimeType: file.mimetype,
          body: fs.createReadStream(file.path),
        },
        fields: 'id, name, webViewLink',
      });
    } catch (driveErr: any) {
      console.error('Error uploading to specific folder, trying root:', driveErr);
      driveResponse = await drive.files.create({
        requestBody: { name: driveFileName },
        media: {
          mimeType: file.mimetype,
          body: fs.createReadStream(file.path),
        },
        fields: 'id, name, webViewLink',
      });
    }

    // 2. Find or Create "AI_WP" Calendar
    let calendarId = 'primary';
    try {
      const calendarList = await calendar.calendarList.list();
      const aiWpCalendar = calendarList.data.items?.find(c => c.summary === 'AI_WP');
      if (aiWpCalendar) {
        calendarId = aiWpCalendar.id!;
      } else {
        const newCalendar = await calendar.calendars.insert({
          requestBody: { summary: 'AI_WP', timeZone: 'Asia/Bangkok' }
        });
        calendarId = newCalendar.data.id!;
      }
    } catch (err: any) {
      console.error('Error finding/creating AI_WP calendar:', err);
      if (err.message?.includes('insufficient authentication scopes')) {
        return res.status(403).json({ 
          error: 'Insufficient permissions for Google Calendar. Please sign out and sign in again to grant required permissions.',
          needsReauth: true
        });
      }
      // Fallback to primary if other error
    }

    const formatDateTime = (dt: string) => {
      if (!dt) return null;
      // Basic validation for YYYY-MM-DD HH:mm or similar
      let formatted = dt.replace(' ', 'T');
      if (!formatted.includes('T')) {
        // If only date is provided, add default time
        formatted += 'T08:00:00';
      }
      const timePart = formatted.split('T')[1];
      if (timePart && timePart.split(':').length === 2) {
        formatted += ':00';
      }
      
      // Final check if it's a valid ISO string
      try {
        const d = new Date(formatted);
        if (isNaN(d.getTime())) return null;
        return formatted;
      } catch {
        return null;
      }
    };

    const spreadsheetId = process.env.GOOGLE_SHEET_ID;
    const processedEvents = [];
    const sheetRows = [];

    // Get initial WP number if sheet exists
    let currentWpNum = 0;
    let currentYearBE = (new Date().getFullYear() + 543) % 100;
    const yearStr = currentYearBE.toString().padStart(2, '0');

    if (spreadsheetId) {
      try {
        const sheetData = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: 'Sheet1!C:C',
        });
        const rows = sheetData.data.values || [];
        const lastWp = rows.length > 0 ? rows[rows.length - 1][0] : null;
        if (lastWp && typeof lastWp === 'string' && lastWp.includes('-')) {
          const [numPart, yearPart] = lastWp.split('-');
          if (parseInt(yearPart, 10) === currentYearBE) {
            currentWpNum = parseInt(numPart, 10);
          }
        }
      } catch (e: any) {
        console.error('Error fetching next WP for multi-process:', e.message);
        if (e.message?.includes('Requested entity was not found')) {
          console.warn('Sheet "Sheet1" not found or spreadsheet ID invalid. Skipping WP auto-increment from sheet.');
        }
      }
    }

    // 3. Loop through events
    for (const eventData of events) {
      currentWpNum++;
      const finalWpNumber = `${currentWpNum.toString().padStart(3, '0')}-${yearStr}`;
      
      const startDT = formatDateTime(eventData.startTime);
      const endDT = formatDateTime(eventData.endTime);

      if (!startDT || !endDT) {
        console.warn(`Invalid date/time for event: ${eventData.stationName}. Skipping calendar entry.`);
        continue;
      }

      // Create Calendar Event
      const calendarEvent = {
        summary: eventData.calendarTitle.replace(eventData.wpNumber, finalWpNumber),
        description: `Automated entry for WP No.${finalWpNumber}\nStation: ${eventData.stationName}\nWork: ${eventData.workDescription}\nUnit: ${eventData.requestingUnit}\nFile: ${driveResponse.data.webViewLink}`,
        start: {
          dateTime: startDT,
          timeZone: 'Asia/Bangkok',
        },
        end: {
          dateTime: endDT,
          timeZone: 'Asia/Bangkok',
        },
      };

      try {
        const calRes = await calendar.events.insert({
          calendarId: calendarId,
          requestBody: calendarEvent,
        });

        processedEvents.push({
          wpNumber: finalWpNumber,
          calendarLink: calRes.data.htmlLink
        });

        // Prepare Sheet Row
        const now = new Date();
        const timestamp = now.toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });
        sheetRows.push([
          timestamp,
          eventData.requestingUnit,
          finalWpNumber,
          eventData.isStaffed ? 'จัดพนักงาน' : 'ไม่จัดพนักงาน',
          eventData.workDescription,
          eventData.startTime,
          eventData.endTime,
          eventData.department
        ]);
      } catch (calErr: any) {
        console.error(`Error creating calendar event for ${eventData.stationName}:`, calErr.message);
        // Continue to next event even if one fails
      }
    }

    // 4. Batch Write to Google Sheet
    let sheetLink = '';
    if (spreadsheetId && sheetRows.length > 0) {
      try {
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: 'Sheet1!A:H',
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: sheetRows }
        });
        sheetLink = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
      } catch (err) {
        console.error('Error writing to Google Sheet:', err);
      }
    }

    // Cleanup local file
    fs.unlinkSync(file.path);

    res.json({
      success: true,
      driveFile: driveResponse.data,
      calendarEvent: { htmlLink: processedEvents[0].calendarLink }, // Return first for UI
      sheetLink: sheetLink,
      processedCount: events.length
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
