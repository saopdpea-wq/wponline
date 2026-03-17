import express from 'express';
// import { createServer as createViteServer } from 'vite'; // Moved to dynamic import
import { google } from 'googleapis';
import multer from 'multer';
import cookieParser from 'cookie-parser';
import path from 'path';
import fs from 'fs';

const app = express();
const PORT = 3000;

// Ensure uploads directory exists (use /tmp for Vercel)
const UPLOAD_DIR = process.env.VERCEL ? '/tmp' : 'uploads';
if (!fs.existsSync(UPLOAD_DIR) && !process.env.VERCEL) {
  fs.mkdirSync(UPLOAD_DIR);
}

const upload = multer({ 
  dest: UPLOAD_DIR,
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

// Helper to get next WP number
async function getNextWpNumber(sheets: any, spreadsheetId: string) {
  let currentWpNum = 0;
  const currentYearBE = (new Date().getFullYear() + 543) % 100;
  const yearStr = currentYearBE.toString().padStart(2, '0');

  try {
    // Dynamically find the first sheet name
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetName = spreadsheet.data.sheets?.[0]?.properties?.title || 'Sheet1';
    
    const sheetData = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:A`,
    });
    const rows = sheetData.data.values || [];
    
    // Search backwards for the last valid WP
    for (let i = rows.length - 1; i >= 0; i--) {
      const val = rows[i][0];
      if (val && typeof val === 'string' && val.includes('-')) {
        const parts = val.split('-');
        if (parts.length === 2) {
          const numPart = parts[0].trim();
          const yearPart = parts[1].trim();
          const num = parseInt(numPart, 10);
          const year = parseInt(yearPart, 10);
          
          if (!isNaN(num) && !isNaN(year)) {
            if (year === currentYearBE) {
              currentWpNum = Math.max(currentWpNum, num);
            }
          }
        }
      }
    }
  } catch (e: any) {
    console.error('Error in getNextWpNumber:', e.message);
  }

  const nextNum = currentWpNum + 1;
  return `${nextNum.toString().padStart(3, '0')}-${yearStr}`;
}

const getAuthClient = async (req?: any) => {
  // 1. Try User OAuth tokens from cookies first (User override)
  const tokens = req?.cookies?.google_tokens;
  if (tokens) {
    try {
      const auth = getOAuth2Client();
      auth.setCredentials(JSON.parse(tokens));
      return auth;
    } catch (e) {
      console.error('Error parsing user tokens:', e);
    }
  }

  // 2. Try Refresh Token from Environment Variables (Always Connected - OAuth method)
  const refreshToken = process.env.GOOGLE_REFRESH_TOKEN;
  if (refreshToken) {
    try {
      const auth = getOAuth2Client();
      auth.setCredentials({ refresh_token: refreshToken });
      return auth;
    } catch (e) {
      console.error('Error using GOOGLE_REFRESH_TOKEN:', e);
    }
  }

  // 3. Fallback to Service Account (Always Connected - Service Account method)
  let serviceAccountJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (serviceAccountJson) {
    try {
      // Clean up the string in case it has extra quotes or escaped characters from env
      serviceAccountJson = serviceAccountJson.trim();
      if (serviceAccountJson.startsWith("'") && serviceAccountJson.endsWith("'")) {
        serviceAccountJson = serviceAccountJson.slice(1, -1);
      }
      if (serviceAccountJson.startsWith('"') && serviceAccountJson.endsWith('"')) {
        try {
          // If it's double-quoted, it might be a JSON-stringified string
          serviceAccountJson = JSON.parse(serviceAccountJson);
        } catch (e) {
          // If parse fails, just use the sliced version
          serviceAccountJson = serviceAccountJson.slice(1, -1);
        }
      }

      const credentials = typeof serviceAccountJson === 'string' ? JSON.parse(serviceAccountJson) : serviceAccountJson;
      const auth = new google.auth.GoogleAuth({
        credentials,
        scopes: SCOPES,
      });
      return await auth.getClient();
    } catch (e: any) {
      console.error('Invalid GOOGLE_SERVICE_ACCOUNT_JSON or error getting client:', e.message);
    }
  }

  return null;
};

// Auth Routes
app.get('/api/auth/url', (req, res) => {
  const isSetup = req.query.setup === 'true';
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent', // Force consent to ensure we get a refresh token
    scope: SCOPES,
    state: isSetup ? 'setup' : undefined
  });
  res.json({ url });
});

app.get('/auth/callback', async (req, res) => {
  const { code, state } = req.query;
  try {
    const { tokens } = await oauth2Client.getToken(code as string);
    
    // Standard login flow
    res.cookie('google_tokens', JSON.stringify(tokens), {
      httpOnly: true,
      secure: true,
      sameSite: 'none',
      maxAge: 365 * 24 * 60 * 60 * 1000 // Increase to 1 year
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
  const serviceAccountJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  const hasRefreshToken = !!process.env.GOOGLE_REFRESH_TOKEN;
  let isServiceAccountValid = false;
  
  if (serviceAccountJson) {
    try {
      let cleaned = serviceAccountJson.trim();
      if (cleaned.startsWith("'") && cleaned.endsWith("'")) cleaned = cleaned.slice(1, -1);
      if (cleaned.startsWith('"') && cleaned.endsWith('"')) {
        try { cleaned = JSON.parse(cleaned); } catch(e) { cleaned = cleaned.slice(1, -1); }
      }
      JSON.parse(typeof cleaned === 'string' ? cleaned : JSON.stringify(cleaned));
      isServiceAccountValid = true;
    } catch (e) {
      isServiceAccountValid = false;
    }
  }
  
  if (tokens) {
    res.json({ isAuthenticated: true, isServiceAccount: false });
  } else if (hasRefreshToken || isServiceAccountValid) {
    res.json({ isAuthenticated: true, isServiceAccount: true });
  } else {
    res.json({ isAuthenticated: false, isServiceAccount: false });
  }
});

app.post('/api/auth/logout', (req, res) => {
  res.clearCookie('google_tokens');
  res.json({ success: true });
});

app.get('/api/auth/token', async (req, res) => {
  const auth = await getAuthClient(req);
  if (!auth) return res.status(401).json({ error: 'Not authenticated' });
  
  try {
    const tokens = await (auth as any).getAccessToken();
    res.json({ token: tokens.token });
  } catch (err: any) {
    console.error('Error getting access token:', err);
    res.status(500).json({ error: 'Failed to get access token' });
  }
});

app.get('/api/next-wp', async (req, res) => {
  const auth = await getAuthClient(req);
  if (!auth) return res.status(401).json({ error: 'Not authenticated' });

  try {
    const sheets = google.sheets({ version: 'v4', auth: auth as any });
    const spreadsheetId = process.env.GOOGLE_SHEET_ID;

    if (!spreadsheetId) return res.json({ nextWp: '001-69' });

    const nextWp = await getNextWpNumber(sheets, spreadsheetId);
    res.json({ nextWp });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch next WP' });
  }
});

// Drive & Calendar Processing
app.post('/api/process', upload.single('file'), async (req: any, res) => {
  const auth = await getAuthClient(req);
  if (!auth) return res.status(401).json({ error: 'Not authenticated' });

  const file = req.file;
  const driveFileId = req.body.driveFileId;
  
  if (!file && !driveFileId) return res.status(400).json({ error: 'No file uploaded or file ID provided' });

  let events = [];
  try {
    events = JSON.parse(req.body.events || '[]');
  } catch (e) {
    return res.status(400).json({ error: 'Invalid events data' });
  }

  if (events.length === 0) return res.status(400).json({ error: 'No events to process' });

  try {
    const drive = google.drive({ version: 'v3', auth: auth as any });
    const calendar = google.calendar({ version: 'v3', auth: auth as any });
    const sheets = google.sheets({ version: 'v4', auth: auth as any });

    // 1. Get next WP number from Google Sheet
    const spreadsheetId = process.env.GOOGLE_SHEET_ID;
    if (!spreadsheetId) throw new Error('GOOGLE_SHEET_ID not set');
    
    const finalWpNumber = await getNextWpNumber(sheets, spreadsheetId);
    console.log(`Generated final WP number: ${finalWpNumber}`);

    // 2. Handle File (Upload or Rename/Move)
    const firstEvent = events[0];
    
    // Clean up station name: remove "สถานีไฟฟ้า" prefix if exists
    const cleanStationName = firstEvent.stationName.replace(/^สถานีไฟฟ้า\s*/, '').trim();
    
    // Clean up requesting unit: remove spaces
    const cleanUnit = (firstEvent.requestingUnit || '').replace(/\s/g, '');
    
    // Clean up date: ensure 2-digit year if it's 4-digit
    let cleanDate = firstEvent.date;
    if (cleanDate.includes('25')) {
      cleanDate = cleanDate.replace(/25(\d{2})/, '$1');
    }

    // Construct shorter filename: WPผจฟ.1No.XXX-XXหน่วยงาน เข้าสถานี (วันที่)งานที่จะทำ.pdf
    // Limit work description to 30 characters
    const shortWorkDesc = (firstEvent.workDescription || '').substring(0, 30);
    const driveFileName = `WPผจฟ.1No.${finalWpNumber}${cleanUnit} เข้า${cleanStationName} (${cleanDate})${shortWorkDesc}.pdf`;
    
    let driveResponse;
    const targetFolderId = process.env.GOOGLE_DRIVE_ROOT_FOLDER_ID || '1aDRqNGcl934p1wzyuFsxd3bdX6PF6CD2';

    if (driveFileId) {
      // File already uploaded by client, just rename and move it
      console.log(`Using existing drive file: ${driveFileId}`);
      try {
        // Update name and move to target folder
        const currentFile = await drive.files.get({ fileId: driveFileId, fields: 'parents' });
        const previousParents = currentFile.data.parents?.join(',') || '';
        
        driveResponse = await drive.files.update({
          fileId: driveFileId,
          addParents: targetFolderId,
          removeParents: previousParents,
          requestBody: {
            name: driveFileName
          },
          fields: 'id, name, webViewLink'
        });
      } catch (updateErr: any) {
        console.error('Error updating existing drive file:', updateErr);
        // If move fails, at least try to rename
        driveResponse = await drive.files.update({
          fileId: driveFileId,
          requestBody: { name: driveFileName },
          fields: 'id, name, webViewLink'
        });
      }
    } else if (file) {
      // Standard upload
      try {
        driveResponse = await drive.files.create({
          requestBody: {
            name: driveFileName,
            parents: [targetFolderId], 
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
    } else {
      throw new Error('No file or file ID provided');
    }

    // 3. Find or Create "AI_WP" Calendar
    let calendarId = '';
    try {
      // Use a larger maxResults to ensure we find it
      const calendarList = await calendar.calendarList.list({ maxResults: 250 });
      const aiWpCalendar = calendarList.data.items?.find(c => c.summary === 'AI_WP');
      
      if (aiWpCalendar) {
        calendarId = aiWpCalendar.id!;
        console.log('Found existing AI_WP calendar:', calendarId);
      } else {
        console.log('AI_WP calendar not found, creating new one...');
        const newCalendar = await calendar.calendars.insert({
          requestBody: { summary: 'AI_WP', timeZone: 'Asia/Bangkok' }
        });
        calendarId = newCalendar.data.id!;
        console.log('Created new AI_WP calendar:', calendarId);
      }
    } catch (err: any) {
      console.error('Error finding/creating AI_WP calendar:', err);
      if (err.message?.includes('insufficient authentication scopes')) {
        return res.status(403).json({ 
          error: 'สิทธิ์การเข้าถึง Google Calendar ไม่เพียงพอ กรุณาออกจากระบบและเข้าใหม่เพื่ออนุญาตสิทธิ์',
          needsReauth: true
        });
      }
    }

    const formatDateTime = (dt: string, isoDateFallback?: string) => {
      if (!dt && !isoDateFallback) return null;
      
      try {
        let datePart = '';
        let timePart = '08:00';

        if (dt) {
          const parts = dt.trim().split(/\s+/);
          if (parts.length >= 2) {
            datePart = parts[0];
            timePart = parts[1];
          } else if (parts.length === 1) {
            if (parts[0].includes('-')) {
              datePart = parts[0];
            } else if (parts[0].includes(':')) {
              datePart = isoDateFallback || '';
              timePart = parts[0];
            } else {
              datePart = isoDateFallback || '';
            }
          }
        } else {
          datePart = isoDateFallback || '';
        }

        if (!datePart) return null;

        // Handle Thai Year (BE) -> AD conversion in datePart
        if (datePart.includes('-')) {
          const dateSegments = datePart.split('-');
          if (dateSegments.length === 3) {
            let year = parseInt(dateSegments[0], 10);
            if (year > 2400) year -= 543;
            datePart = `${year}-${dateSegments[1]}-${dateSegments[2]}`;
          }
        }

        // Ensure timePart has seconds
        const timeSegments = timePart.split(':');
        if (timeSegments.length === 2) {
          timePart = `${timeSegments[0]}:${timeSegments[1]}:00`;
        } else if (timeSegments.length === 1) {
          timePart = `${timeSegments[0]}:00:00`;
        }

        const formatted = `${datePart}T${timePart}`;
        const d = new Date(formatted);
        if (isNaN(d.getTime())) {
          console.error(`Invalid date generated: ${formatted} from input: ${dt}, fallback: ${isoDateFallback}`);
          return null;
        }
        return formatted;
      } catch (e) {
        console.error(`Error formatting date: ${dt}`, e);
        return null;
      }
    };

    // 4. Loop through events
    const processedEvents = [];
    const sheetRows = [];

    const getThaiDate = (date: Date) => {
      const months = [
        'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
        'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'
      ];
      const day = date.getDate();
      const month = months[date.getMonth()];
      const year = (date.getFullYear() + 543) % 100;
      return `${day} ${month} ${year.toString().padStart(2, '0')}`;
    };

    for (const eventData of events) {
      const startDT = formatDateTime(eventData.startTime, eventData.isoDate);
      let endDT = formatDateTime(eventData.endTime, eventData.isoDate);

      if (startDT && !endDT) {
        endDT = startDT;
      }

      // Prepare Sheet Rows (Split by day if multi-day)
      const now = new Date();
      const timestamp = now.toLocaleString('th-TH', { 
        timeZone: 'Asia/Bangkok',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
      });

      if (startDT && endDT) {
        const startDateObj = new Date(startDT);
        const endDateObj = new Date(endDT);
        
        // Extract actual times from the formatted ISO strings (HH:mm)
        const actualStartTime = startDT.split('T')[1].substring(0, 5);
        const actualEndTime = endDT.split('T')[1].substring(0, 5);

        // Create a copy for iteration
        const current = new Date(startDateObj);
        current.setHours(0, 0, 0, 0);
        
        const lastDay = new Date(endDateObj);
        lastDay.setHours(0, 0, 0, 0);

        while (current <= lastDay) {
          const isFirstDay = current.getTime() === (new Date(startDateObj).setHours(0,0,0,0));
          const isLastDay = current.getTime() === lastDay.getTime();
          
          const rowDateISO = current.toISOString().split('T')[0];
          const rowThaiDate = getThaiDate(current);
          
          // Use actual extracted times for the first/last day, otherwise standard 08:00-17:00
          const rowStartTime = isFirstDay ? actualStartTime : '08:00';
          const rowEndTime = isLastDay ? actualEndTime : '17:00';

          sheetRows.push([
            finalWpNumber, // A (WP Number)
            eventData.stationName, // B
            eventData.requestingUnit, // C
            eventData.workDescription, // D
            eventData.isStaffed ? 'จัดพนักงาน' : 'ไม่จัดพนักงาน', // E
            rowThaiDate, // F (Thai format)
            timestamp, // G (Timestamp)
            rowDateISO, // H (Start Date - ISO)
            rowStartTime, // I (Start Time)
            rowDateISO, // J (End Date - ISO)
            rowEndTime, // K (End Time)
            eventData.department || 'ผจฟ.1' // L (Department)
          ]);

          current.setDate(current.getDate() + 1);
        }
      } else {
        // Fallback for cases where dates couldn't be parsed properly
        const actualStartTime = startDT ? startDT.split('T')[1].substring(0, 5) : '08:00';
        const actualEndTime = endDT ? endDT.split('T')[1].substring(0, 5) : '17:00';

        const startDate = startDT ? startDT.split('T')[0] : (eventData.isoDate || '');
        const endDate = endDT ? endDT.split('T')[0] : startDate;

        sheetRows.push([
          finalWpNumber, // A (WP Number)
          eventData.stationName, // B
          eventData.requestingUnit, // C
          eventData.workDescription, // D
          eventData.isStaffed ? 'จัดพนักงาน' : 'ไม่จัดพนักงาน', // E
          eventData.date, // F (Thai format)
          timestamp, // G (Timestamp)
          startDate, // H (Start Date - ISO)
          actualStartTime, // I (Start Time)
          endDate, // J (End Date - ISO)
          actualEndTime, // K (End Time)
          eventData.department || 'ผจฟ.1' // L (Department)
        ]);
      }

      if (!startDT || !endDT || !calendarId) {
        console.warn(`Skipping calendar entry for ${eventData.stationName}. startDT: ${startDT}, endDT: ${endDT}, calendarId: ${calendarId}`);
        continue;
      }

      // Determine colorId based on requesting unit
      // กสฟ. = Blue (9), กดส. = Red (11), Others = Green (10)
      let colorId = '10'; // Default: Green (Basil)
      const unitClean = (eventData.requestingUnit || '').replace(/\s/g, '');
      if (unitClean.includes('กสฟ')) {
        colorId = '9'; // Blue (Blueberry)
      } else if (unitClean.includes('กดส')) {
        colorId = '11'; // Red (Tomato)
      }
      console.log(`Setting colorId ${colorId} for unit: "${eventData.requestingUnit}" (cleaned: "${unitClean}")`);

      // Create Calendar Event
      const calendarEvent = {
        summary: eventData.calendarTitle.replace(eventData.wpNumber, finalWpNumber),
        description: `Automated entry for WP No.${finalWpNumber}\nStation: ${eventData.stationName}\nWork: ${eventData.workDescription}\nUnit: ${eventData.requestingUnit}\nFile: ${driveResponse.data.webViewLink}`,
        colorId: colorId,
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
      } catch (calErr: any) {
        console.error(`Error creating calendar event for ${eventData.stationName}:`, calErr.message);
      }
    }

    // 5. Batch Write to Google Sheet
    let sheetLink = '';
    if (spreadsheetId && sheetRows.length > 0) {
      try {
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: 'Sheet1!A1',
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: sheetRows }
        });
        sheetLink = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
      } catch (err: any) {
        console.error('Error writing to Google Sheet:', err.message);
      }
    }

    // Cleanup local file
    if (file && fs.existsSync(file.path)) {
      fs.unlinkSync(file.path);
    }

    res.json({
      success: true,
      driveFile: driveResponse.data,
      calendarEvent: { htmlLink: processedEvents.length > 0 ? processedEvents[0].calendarLink : null }, // Return first for UI
      sheetLink: sheetLink,
      processedCount: events.length
    });
  } catch (error: any) {
    console.error('Processing error:', error);
    if (file && fs.existsSync(file.path)) {
      try {
        fs.unlinkSync(file.path);
      } catch (e) {
        // Ignore unlink errors in catch block
      }
    }
    res.status(500).json({ error: error.message || 'Internal Server Error' });
  }
});

async function startServer() {
  if (process.env.NODE_ENV !== 'production' && !process.env.VERCEL) {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else if (!process.env.VERCEL) {
    // Standard Express static serving when not on Vercel
    app.use(express.static('dist'));
    app.get('*', (req, res) => {
      res.sendFile(path.resolve('dist/index.html'));
    });
  }

  // Only listen if not on Vercel
  if (!process.env.VERCEL) {
    app.listen(PORT, '0.0.0.0', () => {
      console.log(`Server running on http://localhost:${PORT}`);
    });
  }
}

startServer();

export default app;
