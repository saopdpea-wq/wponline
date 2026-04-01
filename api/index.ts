import express from 'express';
import 'dotenv/config';
// import { createServer as createViteServer } from 'vite'; // Moved to dynamic import
import { google } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
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

// Request Logging Middleware
app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.url}`);
  next();
});

// Google OAuth Configuration
const getRedirectUri = (req?: any) => {
  let redirectUri = '';
  if (process.env.APP_URL) {
    const baseUrl = process.env.APP_URL.replace(/\/$/, '');
    redirectUri = `${baseUrl}/auth/callback`;
  } else if (req) {
    const protocol = req.headers['x-forwarded-proto'] || req.protocol || 'http';
    const host = req.get ? req.get('host') : req.headers?.host;
    if (host) {
      redirectUri = `${protocol}://${host}/auth/callback`;
    }
  }
  
  if (!redirectUri || redirectUri === '/auth/callback') {
    redirectUri = '/auth/callback';
  }

  console.log('Generated Redirect URI:', redirectUri, { 
    hasAppUrl: !!process.env.APP_URL,
    protocol: req?.headers?.['x-forwarded-proto'] || req?.protocol,
    host: req?.get ? req.get('host') : req?.headers?.host
  });
  return redirectUri;
};

const getOAuth2Client = (req?: any) => {
  const clientId = (
    process.env.GOOGLE_CLIENT_ID || 
    process.env.CLIENT_ID || 
    process.env.Google_Client_Id || 
    process.env.google_client_id || 
    ''
  ).trim();
  
  const clientSecret = (
    process.env.GOOGLE_CLIENT_SECRET || 
    process.env.CLIENT_SECRET || 
    process.env.Google_Client_Secret || 
    process.env.google_client_secret || 
    ''
  ).trim();
  
  if (!clientId || !clientSecret) {
    console.error('Missing Google Credentials:', { 
      hasClientId: !!clientId, 
      hasClientSecret: !!clientSecret
    });
    return null;
  }

  const redirectUri = getRedirectUri(req);
  
  // Use the object-based constructor for better compatibility
  return new OAuth2Client({
    clientId,
    clientSecret,
    redirectUri
  });
};

// const oauth2Client = getOAuth2Client(); // Removed to avoid startup errors

const SCOPES = [
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/calendar.events',
  'https://www.googleapis.com/auth/calendar',
  'https://www.googleapis.com/auth/spreadsheets'
];

// Helper to get next WP number
async function getNextWpNumber(sheets: any, spreadsheetId: string, sheetName: string = 'Sheet1') {
  let currentWpNum = 0;
  const currentYearBE = (new Date().getFullYear() + 543) % 100;
  const yearStr = currentYearBE.toString().padStart(2, '0');

  try {
    const sheetData = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:A`,
    });
    const rows = sheetData.data.values || [];
    
    // Search backwards for the last valid WP
    for (let i = rows.length - 1; i >= 0; i--) {
      const row = rows[i];
      if (!row || row.length === 0) continue;
      
      const val = row[0];
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
    console.error(`Error in getNextWpNumber (Sheet: ${sheetName}):`, e.message);
  }

  const nextNum = currentWpNum + 1;
  return `${nextNum.toString().padStart(3, '0')}-${yearStr}`;
}

const getAuthClient = async (req?: any) => {
  // 1. Try User OAuth tokens from cookies first (User override)
  const tokens = req?.cookies?.google_tokens;
  if (tokens) {
    try {
      const auth = getOAuth2Client(req);
      if (auth) {
        auth.setCredentials(JSON.parse(tokens));
        return auth;
      }
    } catch (e) {
      console.error('Error parsing user tokens:', e);
    }
  }

  // 2. Try Refresh Token from Environment Variables (Always Connected - OAuth method)
  const refreshToken = process.env.GOOGLE_REFRESH_TOKEN;
  if (refreshToken) {
    try {
      const auth = getOAuth2Client(req);
      if (auth) {
        auth.setCredentials({ refresh_token: refreshToken });
        return auth;
      }
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
  try {
    const clientId = process.env.GOOGLE_CLIENT_ID || process.env.CLIENT_ID;
    const clientSecret = process.env.GOOGLE_CLIENT_SECRET || process.env.CLIENT_SECRET;

    if (!clientId || !clientSecret) {
      return res.status(400).json({ 
        error: 'กรุณาตั้งค่า GOOGLE_CLIENT_ID และ GOOGLE_CLIENT_SECRET ใน Settings ก่อนใช้งาน' 
      });
    }

    const redirectUri = getRedirectUri(req);
    console.log('Generating auth URL with redirectUri:', redirectUri);
    
    const currentClient = getOAuth2Client(req);
    if (!currentClient) {
      console.error('getOAuth2Client returned null in /api/auth/url');
      return res.status(400).json({ 
        error: 'ไม่สามารถสร้าง OAuth Client ได้ กรุณาตรวจสอบการตั้งค่า GOOGLE_CLIENT_ID และ GOOGLE_CLIENT_SECRET ใน Settings' 
      });
    }
    
    const url = currentClient.generateAuthUrl({
      access_type: 'offline',
      prompt: 'consent', // Force consent to ensure we get a refresh token
      scope: SCOPES,
      state: isSetup ? 'setup' : undefined,
      redirect_uri: redirectUri // Explicitly set it here too
    });
    res.json({ url });
  } catch (error: any) {
    console.error('Error generating auth URL:', error);
    res.status(500).json({ error: 'เกิดข้อผิดพลาดในการสร้าง URL: ' + error.message });
  }
});

app.get('/auth/callback', async (req, res) => {
  const { code, state, error, error_description } = req.query;
  console.log('Auth callback received:', { 
    hasCode: !!code, 
    state, 
    error, 
    error_description,
    fullUrl: req.originalUrl 
  });

  if (error) {
    return res.status(400).send(`
      <html>
        <body>
          <h2>Authentication Error from Google</h2>
          <p style="color: red;">Error: ${error}</p>
          <p>Description: ${error_description}</p>
          <a href="/">กลับหน้าหลัก</a>
        </body>
      </html>
    `);
  }

  if (!code) {
    return res.status(400).send(`
      <html>
        <body>
          <h2>ไม่พบ Authorization Code</h2>
          <p>กรุณากลับไปที่หน้าหลักและลองกดเชื่อมต่อใหม่อีกครั้ง</p>
          <p>หากปัญหายังคงอยู่ กรุณาตรวจสอบว่าได้ตั้งค่า Redirect URI ใน Google Cloud Console ถูกต้องแล้ว</p>
          <a href="/">กลับหน้าหลัก</a>
        </body>
      </html>
    `);
  }

  try {
    const currentClient = getOAuth2Client(req);
    if (!currentClient) {
      console.error('getOAuth2Client failed in callback: Client not configured');
      return res.status(400).send(`
        <html>
          <body>
            <h2>OAuth Client not configured</h2>
            <p>กรุณาตรวจสอบการตั้งค่า GOOGLE_CLIENT_ID และ GOOGLE_CLIENT_SECRET ใน Settings</p>
            <a href="/">กลับหน้าหลัก</a>
          </body>
        </html>
      `);
    }

    console.log('Attempting to exchange code for tokens...');
    const { tokens } = await currentClient.getToken(code as string);
    console.log('Tokens received successfully');
    
    // Standard login flow
    res.cookie('google_tokens', JSON.stringify(tokens), {
      httpOnly: true,
      secure: true,
      sameSite: 'none',
      maxAge: 365 * 24 * 60 * 60 * 1000 // Increase to 1 year
    });

    // Send success message with Refresh Token UI
    const refreshToken = tokens.refresh_token || 'ไม่พบ Refresh Token (อาจเป็นเพราะคุณเคยเชื่อมต่อแล้ว กรุณายกเลิกสิทธิ์ใน Google Account แล้วลองใหม่)';
    
    res.send(`
      <html>
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <script src="https://cdn.tailwindcss.com"></script>
          <link href="https://fonts.googleapis.com/css2?family=Kanit:wght@300;400;600&display=swap" rel="stylesheet">
          <style>
            body { font-family: 'Kanit', sans-serif; background-color: #f8fafc; }
          </style>
        </head>
        <body class="flex items-center justify-center min-h-screen p-4">
          <div class="bg-white p-8 rounded-3xl shadow-xl max-w-2xl w-full border border-slate-100">
            <div class="flex items-center gap-3 mb-6">
              <div class="bg-emerald-500 p-1.5 rounded-lg">
                <svg xmlns="http://www.w3.org/2000/svg" class="h-6 w-6 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="3" d="M5 13l4 4L19 7" />
                </svg>
              </div>
              <h1 class="text-3xl font-bold text-slate-800">คัดลอก <span class="text-indigo-600">Refresh Token</span> ของคุณ</h1>
            </div>

            <p class="text-slate-600 mb-2 text-lg text-left">นำค่าด้านล่างนี้ไปใส่ใน Vercel Environment Variables ชื่อ</p>
            <p class="text-2xl font-bold text-slate-900 mb-6 tracking-tight text-left">GOOGLE_REFRESH_TOKEN</p>

            <div class="relative group mb-8">
              <textarea 
                id="tokenBox" 
                readonly 
                class="w-full h-32 p-5 bg-slate-50 border-2 border-slate-200 rounded-2xl text-slate-700 font-mono text-sm break-all focus:outline-none focus:border-indigo-500 transition-all resize-none"
              >${refreshToken}</textarea>
              <button 
                onclick="copyToken()" 
                class="absolute top-3 right-3 bg-white border border-slate-200 hover:border-indigo-500 text-slate-600 hover:text-indigo-600 px-4 py-2 rounded-xl text-sm font-medium transition-all shadow-sm active:scale-95"
              >
                คัดลอก
              </button>
            </div>

            <div class="bg-rose-50 border border-rose-100 p-5 rounded-2xl">
              <p class="text-rose-600 font-semibold text-center text-lg leading-relaxed">
                ขั้นตอนสุดท้าย: เมื่อใส่ค่าใน Vercel แล้ว อย่าลืมกด <span class="underline decoration-2 underline-offset-4 text-rose-700">Redeploy</span> เพื่อให้ระบบเริ่มทำงานนะครับ
              </p>
            </div>

            <div class="mt-8 text-center flex flex-col gap-4">
              <p class="text-slate-400 text-sm italic">หมายเหตุ: หากคุณต้องการให้แอปทำงานต่อทันทีโดยไม่ต้องปิดหน้าต่างนี้ ระบบได้ส่งสัญญาณไปยังแอปหลักแล้ว</p>
              <button onclick="window.close()" class="text-slate-400 hover:text-slate-600 text-sm transition-all underline">
                ปิดหน้าต่างนี้
              </button>
            </div>
          </div>

          <script>
            function copyToken() {
              const copyText = document.getElementById("tokenBox");
              copyText.select();
              copyText.setSelectionRange(0, 99999);
              navigator.clipboard.writeText(copyText.value);
              
              const btn = event.target;
              const originalText = btn.innerText;
              btn.innerText = "คัดลอกแล้ว!";
              btn.classList.add("bg-indigo-600", "text-white", "border-indigo-600");
              
              setTimeout(() => {
                btn.innerText = originalText;
                btn.classList.remove("bg-indigo-600", "text-white", "border-indigo-600");
              }, 2000);

              if (window.opener) {
                window.opener.postMessage({ type: 'OAUTH_AUTH_SUCCESS' }, '*');
              }
            }
          </script>
        </body>
      </html>
    `);
  } catch (error: any) {
    console.error('Error getting tokens:', error);
    res.status(500).send(`
      <html>
        <body>
          <h2>Authentication failed</h2>
          <p style="color: red;">Error: ${error.message}</p>
          <p>กรุณาตรวจสอบว่าได้ตั้งค่า GOOGLE_CLIENT_ID และ GOOGLE_CLIENT_SECRET ใน Settings ครบถ้วนแล้ว</p>
          <a href="/">กลับหน้าหลัก</a>
        </body>
      </html>
    `);
  }
});

app.get('/api/ping', (req, res) => {
  res.json({ status: 'ok', time: new Date().toISOString() });
});

app.get('/api/test-json', (req, res) => {
  res.json({ message: 'This is a JSON response', status: 200 });
});

app.get('/api/debug-env', (req, res) => {
  const mask = (val: string | undefined) => {
    if (!val) return 'MISSING';
    if (val.length < 8) return 'SET (TOO SHORT)';
    return `${val.substring(0, 4)}...${val.substring(val.length - 4)}`;
  };
  
  res.json({
    GOOGLE_CLIENT_ID: mask(process.env.GOOGLE_CLIENT_ID || process.env.CLIENT_ID),
    GOOGLE_CLIENT_SECRET: mask(process.env.GOOGLE_CLIENT_SECRET || process.env.CLIENT_SECRET),
    GOOGLE_REFRESH_TOKEN: mask(process.env.GOOGLE_REFRESH_TOKEN),
    GOOGLE_SHEET_ID: mask(process.env.GOOGLE_SHEET_ID),
    APP_URL: process.env.APP_URL || 'NOT_SET',
    NODE_ENV: process.env.NODE_ENV,
    VERCEL: process.env.VERCEL || 'NO'
  });
});

app.get('/api/auth/status', (req, res) => {
  const tokens = req.cookies?.google_tokens;
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
    let serviceAccountEmail = null;
    if (isServiceAccountValid && serviceAccountJson) {
      try {
        let cleaned = serviceAccountJson.trim();
        if (cleaned.startsWith("'") && cleaned.endsWith("'")) cleaned = cleaned.slice(1, -1);
        if (cleaned.startsWith('"') && cleaned.endsWith('"')) {
          try { cleaned = JSON.parse(cleaned); } catch(e) { cleaned = cleaned.slice(1, -1); }
        }
        const creds = typeof cleaned === 'string' ? JSON.parse(cleaned) : cleaned;
        serviceAccountEmail = creds.client_email;
      } catch (e) {}
    }
    res.json({ isAuthenticated: true, isServiceAccount: true, serviceAccountEmail });
  } else {
    res.json({ isAuthenticated: false, isServiceAccount: false });
  }
});

app.post('/api/auth/logout', (req, res) => {
  res.clearCookie('google_tokens');
  res.json({ success: true });
});

app.get('/api/auth/token', async (req, res) => {
  try {
    const auth = await getAuthClient(req);
    if (!auth) return res.status(401).json({ error: 'Not authenticated' });
    
    const tokens = await (auth as any).getAccessToken();
    res.json({ token: tokens.token });
  } catch (err: any) {
    console.error('Error getting access token:', err);
    res.status(500).json({ error: 'Failed to get access token: ' + err.message });
  }
});

app.get('/api/next-wp', async (req, res) => {
  try {
    const auth = await getAuthClient(req);
    if (!auth) return res.status(401).json({ error: 'Not authenticated' });

    const sheets = google.sheets({ version: 'v4', auth: auth as any });
    const spreadsheetId = process.env.GOOGLE_SHEET_ID;

    if (!spreadsheetId) return res.json({ nextWp: '001-69' });

    const nextWp = await getNextWpNumber(sheets, spreadsheetId);
    res.json({ nextWp });
  } catch (error: any) {
    console.error('Error in next-wp:', error);
    res.status(500).json({ error: 'Failed to fetch next WP: ' + error.message });
  }
});

// Drive & Calendar Processing
app.post('/api/process', upload.single('file'), async (req: any, res) => {
  const file = req.file;
  try {
    const auth = await getAuthClient(req);
    if (!auth) return res.status(401).json({ error: 'Not authenticated' });

    const driveFileId = req.body.driveFileId;
    
    if (!file && !driveFileId) return res.status(400).json({ error: 'No file uploaded or file ID provided' });

    let events = [];
    try {
      events = JSON.parse(req.body.events || '[]');
    } catch (e) {
      return res.status(400).json({ error: 'Invalid events data' });
    }

    if (events.length === 0) return res.status(400).json({ error: 'No events to process' });

    const drive = google.drive({ version: 'v3', auth: auth as any });
    const calendar = google.calendar({ version: 'v3', auth: auth as any });
    const sheets = google.sheets({ version: 'v4', auth: auth as any });

    // 1. Get next WP number and Sheet Name from Google Sheet
    const spreadsheetId = (process.env.GOOGLE_SHEET_ID || '').trim();
    if (!spreadsheetId) throw new Error('GOOGLE_SHEET_ID not set in environment variables');
    
    // Extract ID if it's a full URL
    let cleanSpreadsheetId = spreadsheetId;
    if (spreadsheetId.includes('/d/')) {
      const match = spreadsheetId.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (match) cleanSpreadsheetId = match[1];
    }

    // Dynamically find the first sheet name
    let sheetName = 'Sheet1';
    let finalWpNumber = '001-69';
    try {
      const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: cleanSpreadsheetId });
      sheetName = spreadsheet.data.sheets?.[0]?.properties?.title || 'Sheet1';
      
      // Get next WP number using the found sheet name
      finalWpNumber = await getNextWpNumber(sheets, cleanSpreadsheetId, sheetName);
    } catch (sheetErr: any) {
      console.error('Error getting spreadsheet info:', sheetErr.message);
      // If we can't even get the spreadsheet info, it's likely a permission or ID issue
      if (sheetErr.code === 404) throw new Error(`Spreadsheet not found. Please check GOOGLE_SHEET_ID: ${cleanSpreadsheetId}`);
      if (sheetErr.code === 403) throw new Error(`Permission denied for Spreadsheet. Please share it with the Service Account or ensure you are logged in with the right account.`);
      
      // Fallback to default if it's some other error
      finalWpNumber = await getNextWpNumber(sheets, cleanSpreadsheetId, 'Sheet1');
    }

    console.log(`Generated final WP number: ${finalWpNumber} using sheet: ${sheetName}`);

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
    
    let driveResponse: any = null;
    let driveFileLink = '';
    
    // Clean up target folder ID if it's a URL
    const rawFolderId = (process.env.GOOGLE_DRIVE_ROOT_FOLDER_ID || '').trim();
    let targetFolderId = 'root';
    if (rawFolderId) {
      targetFolderId = rawFolderId;
      if (rawFolderId.includes('/folders/')) {
        const match = rawFolderId.match(/\/folders\/([a-zA-Z0-9-_]+)/);
        if (match) targetFolderId = match[1];
      }
    }

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
        driveFileLink = driveResponse.data.webViewLink || '';
      } catch (updateErr: any) {
        console.error('Error updating existing drive file:', updateErr.message);
        // If move fails, at least try to rename
        try {
          driveResponse = await drive.files.update({
            fileId: driveFileId,
            requestBody: { name: driveFileName },
            fields: 'id, name, webViewLink'
          });
          driveFileLink = driveResponse.data.webViewLink || '';
        } catch (renameErr: any) {
          console.error('Final attempt to rename failed:', renameErr.message);
          // If we have the ID, we can still construct a link even if update failed
          driveFileLink = `https://drive.google.com/file/d/${driveFileId}/view`;
        }
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
        driveFileLink = driveResponse.data.webViewLink || '';
      } catch (driveErr: any) {
        console.error('Error uploading to specific folder, trying root:', driveErr.message);
        try {
          driveResponse = await drive.files.create({
            requestBody: { name: driveFileName },
            media: {
              mimeType: file.mimetype,
              body: fs.createReadStream(file.path),
            },
            fields: 'id, name, webViewLink',
          });
          driveFileLink = driveResponse.data.webViewLink || '';
        } catch (rootErr: any) {
          console.error('Upload to root also failed:', rootErr.message);
          throw new Error(`Failed to upload file to Google Drive: ${rootErr.message}`);
        }
      }
    } else {
      throw new Error('No file or file ID provided');
    }

    console.log(`Drive file handled. Link: ${driveFileLink}`);

    // 3. Find or Create "AI_WP" Calendar
    let calendarId = '';
    let calendarError = null;
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
      console.error('Error finding/creating AI_WP calendar:', err.message);
      calendarError = `ไม่สามารถเข้าถึงหรือสร้างปฏิทิน "AI_WP" ได้: ${err.message}`;
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
            eventData.department || 'ผจฟ.1', // L (Department)
            driveFileLink // M (Drive Link)
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
          eventData.department || 'ผจฟ.1', // L (Department)
          driveFileLink // M (Drive Link)
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
        description: `Automated entry for WP No.${finalWpNumber}\nStation: ${eventData.stationName}\nWork: ${eventData.workDescription}\nUnit: ${eventData.requestingUnit}\nFile: ${driveFileLink}`,
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
    let sheetError = null;
    if (cleanSpreadsheetId && sheetRows.length > 0) {
      try {
        await sheets.spreadsheets.values.append({
          spreadsheetId: cleanSpreadsheetId,
          range: `${sheetName}!A1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: sheetRows }
        });
        sheetLink = `https://docs.google.com/spreadsheets/d/${cleanSpreadsheetId}`;
      } catch (err: any) {
        console.error(`Error writing to Google Sheet (${sheetName}):`, err.message);
        sheetError = err.message;
        // Try fallback to Sheet1 if dynamic name failed
        if (sheetName !== 'Sheet1') {
          try {
            await sheets.spreadsheets.values.append({
              spreadsheetId: cleanSpreadsheetId,
              range: 'Sheet1!A1',
              valueInputOption: 'USER_ENTERED',
              requestBody: { values: sheetRows }
            });
            sheetLink = `https://docs.google.com/spreadsheets/d/${cleanSpreadsheetId}`;
            sheetError = null; // Fallback succeeded
          } catch (fallbackErr: any) {
            console.error('Fallback Sheet1 write also failed:', fallbackErr.message);
            sheetError = fallbackErr.message;
          }
        }
      }
    }

    // Cleanup local file
    if (file && fs.existsSync(file.path)) {
      fs.unlinkSync(file.path);
    }

    res.json({
      success: true,
      driveFile: driveResponse?.data || { name: driveFileName, id: driveFileId, webViewLink: driveFileLink },
      calendarEvent: { htmlLink: processedEvents.length > 0 ? processedEvents[0].calendarLink : null },
      sheetLink: sheetLink,
      sheetError: sheetError,
      calendarError: calendarError,
      processedCount: events.length,
      calendarCount: processedEvents.length
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

app.post('/api/extract-pdf', upload.single('file'), async (req: any, res) => {
  const file = req.file;
  if (!file) return res.status(400).json({ error: 'No file uploaded' });

  try {
    console.log('PDF Extraction started for file:', file.originalname, 'size:', file.size);
    const dataBuffer = fs.readFileSync(file.path);
    console.log('File read into buffer, length:', dataBuffer.length);
    
    // Fix for "DOMMatrix is not defined" error in Node.js
    // We should use a more robust mock or a real polyfill if possible
    if (typeof (global as any).DOMMatrix === 'undefined') {
      try {
        // Try to get DOMMatrix from @napi-rs/canvas which is a dependency of pdf-parse
        const canvas = await import('@napi-rs/canvas');
        (global as any).DOMMatrix = (canvas as any).DOMMatrix;
        console.log('DOMMatrix polyfilled from @napi-rs/canvas');
      } catch (e) {
        // Fallback to a slightly better mock if canvas import fails
        console.warn('Failed to import @napi-rs/canvas for DOMMatrix, using mock');
        (global as any).DOMMatrix = class DOMMatrix {
          m11 = 1; m12 = 0; m13 = 0; m14 = 0;
          m21 = 0; m22 = 1; m23 = 0; m24 = 0;
          m31 = 0; m32 = 0; m33 = 1; m34 = 0;
          m41 = 0; m42 = 0; m43 = 0; m44 = 1;
          constructor() {}
          static fromFloat32Array() { return new DOMMatrix(); }
          static fromFloat64Array() { return new DOMMatrix(); }
          multiply() { return this; }
          inverse() { return this; }
          translate() { return this; }
          scale() { return this; }
          rotate() { return this; }
        };
      }
    }

    console.log('Starting PDF extraction process...');
    let extractedText = '';
    
    try {
      // Try using pdf-parse first
      console.log('Attempting to use pdf-parse...');
      
      const pdfModule: any = await import('pdf-parse');
      
      // Check if it's the new version (Mehmet Kozan's) or the old version
      if (pdfModule.PDFParse) {
        console.log('Using new PDFParse API...');
        
        // Fix for "Setting up fake worker failed" error
        try {
          let workerUrl = '';
          
          const path = await import('path');
          const fs = await import('fs');
          const { pathToFileURL, fileURLToPath } = await import('url');
          
          // Get current directory
          const __filename = fileURLToPath(import.meta.url);
          const __dirname = path.dirname(__filename);
          
          // Try to find local worker file in multiple possible locations
          const possiblePaths = [
            // Relative to current file
            path.resolve(__dirname, '..', 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.mjs'),
            path.resolve(__dirname, '..', 'node_modules', 'pdfjs-dist', 'build', 'pdf.worker.mjs'),
            // Absolute from process root
            path.join(process.cwd(), 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.mjs'),
            path.join(process.cwd(), 'node_modules', 'pdfjs-dist', 'build', 'pdf.worker.mjs'),
            // Vercel specific
            '/var/task/node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs',
            '/var/task/node_modules/pdfjs-dist/build/pdf.worker.mjs'
          ];
          
          for (const p of possiblePaths) {
            try {
              if (fs.existsSync(p)) {
                workerUrl = pathToFileURL(p).href;
                console.log('Found PDF worker at:', p);
                console.log('Converted to URL:', workerUrl);
                break;
              }
            } catch (e) {}
          }
          
          if (workerUrl) {
            console.log('Setting PDF worker to:', workerUrl);
            try {
              pdfModule.PDFParse.setWorker(workerUrl);
            } catch (setWorkerError) {
              console.warn('pdfModule.PDFParse.setWorker failed, continuing...', setWorkerError);
            }
          } else {
            console.warn('Could not find pdf.worker.mjs, pdf-parse will attempt to use fake worker');
          }
        } catch (workerSetupError) {
          console.warn('Error during PDF worker setup:', workerSetupError);
        }

        const parser = new pdfModule.PDFParse({ 
          data: dataBuffer,
          disableWorker: true, // Still use disableWorker for better stability in serverless
          verbosity: 0
        });
        const result = await parser.getText();
        extractedText = result.text;
        await parser.destroy();
      } else {
        // Old version or default export is a function
        const pdf = pdfModule.default || (typeof pdfModule === 'function' ? pdfModule : null);
        if (typeof pdf === 'function') {
          console.log('Using legacy pdf-parse function API...');
          const data = await pdf(dataBuffer);
          extractedText = typeof data === 'string' ? data : (data as any).text || '';
        } else {
          throw new Error('PDF parser function/class not found in module');
        }
      }
      
      console.log('PDF parsed successfully, length:', extractedText.length);
    } catch (pdfError: any) {
      console.warn('pdf-parse failed or could not be initialized, falling back to Gemini AI...', pdfError.message);
      
      // Fallback to Gemini AI for PDF extraction
      if (process.env.GEMINI_API_KEY) {
        try {
          const { GoogleGenAI } = await import('@google/genai');
          const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
          
          const result = await ai.models.generateContent({
            model: "gemini-3-flash-preview",
            contents: [{
              parts: [
                { text: "Extract all text from this PDF document accurately." },
                { inlineData: { data: dataBuffer.toString('base64'), mimeType: "application/pdf" } }
              ]
            }]
          });
          
          extractedText = result.text || '';
          console.log('PDF extracted successfully using Gemini AI fallback, length:', extractedText.length);
        } catch (aiError: any) {
          console.error('Gemini AI fallback also failed:', aiError.message);
          throw new Error('ทั้งระบบสกัดข้อความปกติและ AI ไม่สามารถอ่านไฟล์นี้ได้: ' + aiError.message);
        }
      } else {
        throw new Error('ระบบสกัดข้อความปกติขัดข้อง และไม่มี AI Key สำหรับระบบสำรอง: ' + pdfError.message);
      }
    }
    
    // Cleanup local file
    if (fs.existsSync(file.path)) {
      fs.unlinkSync(file.path);
    }

    if (!extractedText || extractedText.trim().length === 0) {
      throw new Error('ไม่พบข้อความในไฟล์ PDF นี้ หรือไฟล์อาจเป็นรูปภาพที่ไม่มีข้อความ');
    }

    res.json({ text: extractedText });
  } catch (error: any) {
    console.error('PDF extraction error details:', {
      message: error.message,
      stack: error.stack,
      file: file.originalname
    });
    if (file && fs.existsSync(file.path)) {
      try {
        fs.unlinkSync(file.path);
      } catch (e) {}
    }
    res.status(500).json({ error: 'Failed to extract text from PDF: ' + error.message });
  }
});

app.get('/api/diag', (req, res) => {
  const getEnvStatus = (key: string) => {
    const val = process.env[key] || process.env[key.toLowerCase()] || process.env[key.charAt(0).toUpperCase() + key.slice(1).toLowerCase()];
    return val ? 'SET' : 'MISSING';
  };

  res.json({
    APP_URL: process.env.APP_URL || 'NOT_SET',
    GOOGLE_CLIENT_ID: getEnvStatus('GOOGLE_CLIENT_ID'),
    GOOGLE_CLIENT_SECRET: getEnvStatus('GOOGLE_CLIENT_SECRET'),
    GOOGLE_SHEET_ID: getEnvStatus('GOOGLE_SHEET_ID'),
    GOOGLE_DRIVE_ROOT_FOLDER_ID: getEnvStatus('GOOGLE_DRIVE_ROOT_FOLDER_ID'),
    GEMINI_API_KEY: getEnvStatus('GEMINI_API_KEY'),
    NODE_ENV: process.env.NODE_ENV,
    VERCEL: process.env.VERCEL || 'NO'
  });
});

// Global Error Handler
app.use('/api/*', (req, res, next) => {
  res.status(404).json({ error: `API route not found: ${req.originalUrl}` });
});

app.use((err: any, req: any, res: any, next: any) => {
  console.error('Global Error Handler:', err);
  const status = err.status || 500;
  const message = err.message || 'Internal Server Error';
  
  if (req.path.startsWith('/api')) {
    res.status(status).json({ error: message, code: err.code });
  } else {
    res.status(status).send(`
      <html>
        <body>
          <h2>Server Error</h2>
          <p style="color: red;">${message}</p>
          <a href="/">กลับหน้าหลัก</a>
        </body>
      </html>
    `);
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
