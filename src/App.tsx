import React, { useState, useEffect, useRef } from 'react';
import { GoogleGenAI, Type } from "@google/genai";
import { 
  Upload, 
  FileText, 
  Calendar, 
  CheckCircle2, 
  AlertCircle, 
  Loader2, 
  LogOut, 
  ExternalLink,
  RefreshCw,
  FileCheck,
  Hash,
  User,
  Clock,
  ChevronRight,
  Table
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

// Initialize Gemini safely
const getGeminiKey = () => {
  try {
    return process.env.GEMINI_API_KEY || '';
  } catch (e) {
    return '';
  }
};

const ai = new GoogleGenAI({ apiKey: getGeminiKey() });

interface ExtractedEvent {
  wpNumber: string;
  stationName: string;
  date: string;
  isoDate: string;
  calendarTitle: string;
  isStaffed: boolean;
  requestingUnit: string;
  workDescription: string;
  startTime: string;
  endTime: string;
  department: string;
}

interface ExtractedData {
  events: ExtractedEvent[];
}

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState<boolean | null>(null);
  const [isServiceAccount, setIsServiceAccount] = useState<boolean>(false);
  const [serviceAccountEmail, setServiceAccountEmail] = useState<string | null>(null);
  const [file, setFile] = useState<File | null>(null);
  const [isExtracting, setIsExtracting] = useState(false);
  const [extractedData, setExtractedData] = useState<ExtractedData | null>(null);
  const [nextWp, setNextWp] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [result, setResult] = useState<{ 
    driveLink: string; 
    calendarLink: string; 
    sheetLink?: string; 
    sheetError?: string;
    calendarError?: string;
    processedCount?: number;
    calendarCount?: number;
  } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    checkAuthStatus();
    
    const handleMessage = (event: MessageEvent) => {
      // Validate origin is from AI Studio preview or localhost
      const origin = event.origin;
      if (!origin.endsWith('.run.app') && !origin.includes('localhost')) {
        return;
      }
      
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        setIsAuthenticated(true);
        fetchNextWp();
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

  useEffect(() => {
    // Auto-login attempt if not authenticated and haven't tried this session
    if (isAuthenticated === false && !sessionStorage.getItem('auto_login_attempted')) {
      sessionStorage.setItem('auto_login_attempted', 'true');
      // Small delay to let the UI render first
      setTimeout(() => {
        handleLogin(true); // true means it's an auto-attempt
      }, 1000);
    }
  }, [isAuthenticated]);

  const fetchNextWp = async () => {
    try {
      const res = await fetch('/api/next-wp');
      if (!res.ok) return;
      const contentType = res.headers.get('content-type');
      if (contentType && contentType.includes('application/json')) {
        const data = await res.json();
        if (data.nextWp) setNextWp(data.nextWp);
      }
    } catch (err) {
      console.error('Failed to fetch next WP', err);
    }
  };

  const checkAuthStatus = async () => {
    try {
      const res = await fetch('/api/auth/status');
      const contentType = res.headers.get('content-type');
      if (contentType && contentType.includes('application/json')) {
        const data = await res.json();
        setIsAuthenticated(data.isAuthenticated);
        setIsServiceAccount(!!data.isServiceAccount);
        setServiceAccountEmail(data.serviceAccountEmail);
        if (data.isAuthenticated) fetchNextWp();
      } else {
        const text = await res.text();
        console.error('Non-JSON response from /api/auth/status:', text);
        setIsAuthenticated(false);
      }
    } catch (err) {
      console.error('Auth check failed', err);
      setIsAuthenticated(false);
    }
  };

  const handleLogin = async (isAuto = false, isSetup = false) => {
    try {
      const res = await fetch(`/api/auth/url${isSetup ? '?setup=true' : ''}`);
      
      let data;
      const contentType = res.headers.get('content-type');
      if (contentType && contentType.includes('application/json')) {
        data = await res.json();
      } else {
        const text = await res.text();
        console.error('Non-JSON response from server:', text);
        if (!isAuto) setError('Server ส่งข้อมูลกลับมาไม่ถูกต้อง (ไม่ใช่ JSON)');
        return;
      }
      
      if (!res.ok) {
        if (!isAuto) setError(data.error || 'ไม่สามารถดึง URL สำหรับเชื่อมต่อได้');
        return;
      }

      const { url } = data;
      const popup = window.open(url, 'google_oauth', 'width=600,height=700');
      
      if (!popup || popup.closed || typeof popup.closed === 'undefined') {
        console.warn('Popup blocked');
        if (!isAuto) {
          setError('เบราว์เซอร์บล็อกหน้าต่างป๊อปอัพ กรุณาอนุญาตป๊อปอัพแล้วลองใหม่อีกครั้ง');
        }
      }
    } catch (err) {
      setError('ไม่สามารถเชื่อมต่อกับ Server ได้');
    }
  };

  const handleLogout = async () => {
    await fetch('/api/auth/logout', { method: 'POST' });
    setIsAuthenticated(false);
    resetState();
  };

  const resetState = () => {
    setFile(null);
    setExtractedData(null);
    setResult(null);
    setError(null);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile) {
      setFile(selectedFile);
      extractData(selectedFile);
    }
  };

  const extractData = async (file: File) => {
    setIsExtracting(true);
    setError(null);
    setExtractedData(null);

    try {
      let contents: any[] = [];
      
      if (file.type === 'application/pdf') {
        // Use server-side extraction for PDF to get text first
        // This is more reliable for text-based PDFs
        const formData = new FormData();
        formData.append('file', file);
        
        const extractRes = await fetch('/api/extract-pdf', {
          method: 'POST',
          body: formData
        });
        
        if (!extractRes.ok) {
          throw new Error('ไม่สามารถสกัดข้อความจาก PDF ได้');
        }
        
        const { text } = await extractRes.json();
        contents = [
          {
            text: `Extract information from this Work Permit (WP) text for Thai electrical substations.
            
            TEXT FROM PDF:
            ${text}
            
            IMPORTANT: If the document contains multiple work entries, they MUST ALL share the same WP Number found in the document. Do not generate different WP numbers for different entries within the same file.
            
            For EACH entry found, extract:
            1. WP Number (e.g., 013-69)
            2. Station Name (e.g., สถานีไฟฟ้ากระทุ่มแบน 6)
            3. Date of work (Thai format, e.g., 13 ก.พ. 69)
            4. ISO Date (YYYY-MM-DD)
            5. Is it staffed (จัดพนักงาน) or unstaffed (ไม่จัดพนักงาน)? Look for keywords like "จัดพนักงาน" or "ไม่จัดพนักงาน". Default to "จัดพนักงาน" if unsure.
            6. Requesting Unit (หน่วยงานที่ขออนุญาตทำงาน): Extract the exact name of the department or unit requesting the permit. This is a critical field. Look for it near "หน่วยงานที่ขออนุญาต", "สังกัด", or "ผู้ขออนุญาต" (e.g., กฟภ. ผคส.กสฟ.(ก3)). Ensure the name is captured accurately without extra characters.
            7. Work Description (งานที่จะทำ)
            8. Start Date/Time (วันเวลาที่ขออนุญาตเริ่มต้น, format: YYYY-MM-DD HH:mm)
            9. End Date/Time (วันเวลาที่ขออนุญาตสิ้นสุด, format: YYYY-MM-DD HH:mm)
            10. Department (แผนก, default to "ผจฟ.1" if not found)
            
            Generate a Calendar Title pattern for each: "[Station Name] ([Staffed Status]) WP ผจฟ.1 No.[WP Number] บำรุงรักษาระบบ SCPS ประจำปี (ผปค.กสฟ.ก3)"
            
            Return a JSON object with an "events" array containing all found entries.`
          }
        ];
      } else {
        // For images, send directly to Gemini
        const reader = new FileReader();
        const fileDataPromise = new Promise<string>((resolve) => {
          reader.onload = () => resolve((reader.result as string).split(',')[1]);
          reader.readAsDataURL(file);
        });
        const base64Data = await fileDataPromise;
        
        contents = [
          {
            parts: [
              {
                inlineData: {
                  mimeType: file.type || 'image/jpeg',
                  data: base64Data
                }
              },
              {
                text: `Extract information from this Work Permit (WP) document for Thai electrical substations.
                The document may contain multiple pages, and each page or section might represent a different work entry.
                
                IMPORTANT: If the document contains multiple work entries, they MUST ALL share the same WP Number found in the document. Do not generate different WP numbers for different entries within the same file.
                
                For EACH entry found, extract:
                1. WP Number (e.g., 013-69)
                2. Station Name (e.g., สถานีไฟฟ้ากระทุ่มแบน 6)
                3. Date of work (Thai format, e.g., 13 ก.พ. 69)
                4. ISO Date (YYYY-MM-DD)
                5. Is it staffed (จัดพนักงาน) or unstaffed (ไม่จัดพนักงาน)? Look for keywords like "จัดพนักงาน" or "ไม่จัดพนักงาน". Default to "จัดพนักงาน" if unsure.
                6. Requesting Unit (หน่วยงานที่ขออนุญาตทำงาน): Extract the exact name of the department or unit requesting the permit. This is a critical field. Look for it near "หน่วยงานที่ขออนุญาต", "สังกัด", or "ผู้ขออนุญาต" (e.g., กฟภ. ผคส.กสฟ.(ก3)). Ensure the name is captured accurately without extra characters.
                7. Work Description (งานที่จะทำ)
                8. Start Date/Time (วันเวลาที่ขออนุญาตเริ่มต้น, format: YYYY-MM-DD HH:mm)
                9. End Date/Time (วันเวลาที่ขออนุญาตสิ้นสุด, format: YYYY-MM-DD HH:mm)
                10. Department (แผนก, default to "ผจฟ.1" if not found)
                
                Generate a Calendar Title pattern for each: "[Station Name] ([Staffed Status]) WP ผจฟ.1 No.[WP Number] บำรุงรักษาระบบ SCPS ประจำปี (ผปค.กสฟ.ก3)"
                
                Return a JSON object with an "events" array containing all found entries.`
              }
            ]
          }
        ];
      }

      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: contents,
        config: {
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              events: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    wpNumber: { type: Type.STRING },
                    stationName: { type: Type.STRING },
                    date: { type: Type.STRING },
                    isoDate: { type: Type.STRING },
                    isStaffed: { type: Type.BOOLEAN },
                    calendarTitle: { type: Type.STRING },
                    requestingUnit: { type: Type.STRING },
                    workDescription: { type: Type.STRING },
                    startTime: { type: Type.STRING },
                    endTime: { type: Type.STRING },
                    department: { type: Type.STRING }
                  },
                  required: ["wpNumber", "stationName", "date", "isoDate", "isStaffed", "calendarTitle", "requestingUnit", "workDescription", "startTime", "endTime", "department"]
                }
              }
            },
            required: ["events"]
          }
        }
      });

      const data = JSON.parse(response.text || '{}');
      
      // Post-process to ensure same WP number if multiple events
      if (data.events && data.events.length > 1) {
        const firstWp = data.events[0].wpNumber;
        data.events = data.events.map((ev: any) => ({
          ...ev,
          wpNumber: firstWp
        }));
      }
      
      setExtractedData(data);
    } catch (err: any) {
      console.error('Extraction failed', err);
      if (err.message?.includes('API_KEY_INVALID') || err.message?.includes('API key not found')) {
        setError('API Key สำหรับ Gemini ไม่ถูกต้องหรือไม่ได้ตั้งค่า กรุณาตรวจสอบการตั้งค่า');
      } else if (err.message?.includes('SAFETY')) {
        setError('ไม่สามารถสกัดข้อมูลได้เนื่องจากติดตัวกรองความปลอดภัยของ AI กรุณาลองใช้ไฟล์อื่น');
      } else {
        setError(err.message || 'ไม่สามารถสกัดข้อมูลจากไฟล์ได้ กรุณาลองใหม่อีกครั้งหรือตรวจสอบรูปแบบไฟล์');
      }
    } finally {
      setIsExtracting(false);
    }
  };

  const handleProcess = async () => {
    if (!file || !extractedData || extractedData.events.length === 0) return;
    setIsProcessing(true);
    setError(null);

    try {
      let driveFileId = null;
      
      // If file is large (> 4MB), upload to Drive directly from client
      // to bypass server payload limits (FUNCTION_PAYLOAD_TOO_LARGE)
      if (file.size > 4 * 1024 * 1024) {
        console.log('File is large, uploading directly to Drive from client...');
        try {
          const tokenRes = await fetch('/api/auth/token');
          if (!tokenRes.ok) throw new Error('Failed to get auth token for direct upload');
          const { token } = await tokenRes.json();
          
          const boundary = '-------314159265358979323846';
          const delimiter = "\r\n--" + boundary + "\r\n";
          const close_delim = "\r\n--" + boundary + "--";

          const metadata = {
            name: file.name,
            mimeType: file.type
          };

          const metadataPart = delimiter +
            'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
            JSON.stringify(metadata) +
            delimiter +
            'Content-Type: ' + file.type + '\r\n\r\n';

          const multipartBody = new Blob([metadataPart, file, close_delim], { type: 'multipart/related; boundary=' + boundary });

          const uploadRes = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
            method: 'POST',
            headers: {
              'Authorization': `Bearer ${token}`
            },
            body: multipartBody
          });
          
          if (!uploadRes.ok) {
            const errText = await uploadRes.text();
            throw new Error(`Direct upload failed: ${errText}`);
          }
          
          const uploadData = await uploadRes.json();
          driveFileId = uploadData.id;
          console.log('Direct upload success, fileId:', driveFileId);
        } catch (uploadErr: any) {
          console.error('Direct upload error:', uploadErr);
          // Fallback to standard upload if direct fails (maybe it's just under the limit)
        }
      }

      const formData = new FormData();
      if (driveFileId) {
        formData.append('driveFileId', driveFileId);
      } else {
        formData.append('file', file);
      }
      formData.append('events', JSON.stringify(extractedData.events));

      const res = await fetch('/api/process', {
        method: 'POST',
        body: formData
      });
      
      if (!res.ok) {
        const errorText = await res.text();
        let errorMessage = `Server error: ${res.status}`;
        try {
          const errorJson = JSON.parse(errorText);
          errorMessage = errorJson.error || errorMessage;
        } catch {
          errorMessage = `${errorMessage}. ${errorText.substring(0, 100)}`;
        }

        setError(errorMessage);
        
        if (res.status === 401) {
          setIsAuthenticated(false);
        }
        return;
      }

      const contentType = res.headers.get('content-type');
      if (!contentType || !contentType.includes('application/json')) {
        const text = await res.text();
        throw new Error(`Server returned non-JSON response: ${text.substring(0, 50)}...`);
      }

      const data = await res.json();
      if (data.success) {
        setResult({
          driveLink: data.driveFile.webViewLink,
          calendarLink: data.calendarEvent.htmlLink,
          sheetLink: data.sheetLink,
          sheetError: data.sheetError,
          calendarError: data.calendarError,
          processedCount: data.processedCount,
          calendarCount: data.calendarCount
        });
        fetchNextWp(); // Refresh for the next upload
      } else {
        if (res.status === 403 && data.needsReauth) {
          setError(data.error);
          setIsAuthenticated(false); // Force re-login button to show
        } else {
          setError(data.error || 'Processing failed');
        }
      }
    } catch (err: any) {
      console.error('Connection error:', err);
      setError(`Connection error: ${err.message || 'Please try again.'}`);
    } finally {
      setIsProcessing(false);
    }
  };

  if (isAuthenticated === null) {
    return (
      <div className="min-h-screen bg-stone-50 flex items-center justify-center">
        <Loader2 className="w-8 h-8 animate-spin text-stone-400" />
      </div>
    );
  }

  if (isAuthenticated === false) {
    return (
      <div className="min-h-screen bg-stone-50 flex items-center justify-center p-6">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="max-w-md w-full bg-white rounded-[2.5rem] p-12 shadow-xl border border-stone-100 text-center"
        >
          <div className="w-20 h-20 bg-emerald-600 rounded-3xl flex items-center justify-center text-white mx-auto mb-8 shadow-lg shadow-emerald-200">
            <Calendar className="w-10 h-10" />
          </div>
          <h1 className="text-3xl font-bold text-stone-900 mb-4 tracking-tight">ยินดีต้อนรับ</h1>
          <p className="text-stone-500 mb-10 leading-relaxed">
            กรุณาเชื่อมต่อบัญชี Google ของคุณเพื่อเข้าใช้งานระบบจัดการ WP อัตโนมัติ
          </p>
          
          <button 
            onClick={() => handleLogin()}
            className="w-full flex items-center justify-center gap-3 text-lg font-bold text-white bg-stone-900 hover:bg-stone-800 transition-all py-4 rounded-2xl shadow-lg active:scale-[0.98]"
          >
            <Calendar className="w-6 h-6" />
            เชื่อมต่อ Google
          </button>
          
          {error && (
            <p className="mt-6 text-sm text-red-500 font-medium bg-red-50 py-3 px-4 rounded-xl border border-red-100">
              {error}
            </p>
          )}
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-stone-50 text-stone-900 font-sans selection:bg-emerald-100">
      {/* Header */}
      <header className="border-b border-stone-200 bg-white/80 backdrop-blur-md sticky top-0 z-10">
        <div className="max-w-5xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-emerald-600 rounded-lg flex items-center justify-center text-white">
              <FileCheck className="w-5 h-5" />
            </div>
            <h1 className="font-semibold text-lg tracking-tight">ระบบจัดการ WP อัตโนมัติ</h1>
          </div>
          <div className="flex items-center gap-4">
            {isAuthenticated ? (
              <div className="flex items-center gap-3">
                <span className="flex items-center gap-1.5 text-xs font-bold text-emerald-600 bg-emerald-50 px-2 py-1 rounded-full uppercase tracking-wider">
                  <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                  เชื่อมต่อแล้ว
                </span>
                {!isServiceAccount && (
                  <button 
                    onClick={handleLogout}
                    className="p-2 text-stone-400 hover:text-stone-900 transition-colors rounded-lg hover:bg-stone-100"
                    title="ออกจากระบบ"
                  >
                    <LogOut className="w-4 h-4" />
                  </button>
                )}
              </div>
            ) : (
              <button 
                onClick={() => handleLogin()}
                className="flex items-center gap-2 text-sm font-bold text-stone-600 hover:text-stone-900 transition-colors bg-white border border-stone-200 px-4 py-2 rounded-xl shadow-sm hover:shadow-md active:scale-95"
              >
                <Calendar className="w-4 h-4" />
                เชื่อมต่อ Google
              </button>
            )}
          </div>
        </div>
      </header>

      <main className="max-w-3xl mx-auto px-6 py-12">
        <div className="space-y-8">
          {/* Next WP Indicator */}
          {isAuthenticated && nextWp && (
            <div className="bg-white border border-stone-200 rounded-2xl p-4 flex items-center justify-between shadow-sm">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-stone-100 rounded-xl flex items-center justify-center text-stone-500">
                  <Hash className="w-5 h-5" />
                </div>
                <div>
                  <p className="text-xs font-bold text-stone-400 uppercase tracking-wider">เลขที่ WP ถัดไป</p>
                  <p className="font-mono font-bold text-stone-900">{nextWp}</p>
                </div>
              </div>
            </div>
          )}

          {/* Upload Section */}
          <section>
            <div 
              onClick={() => fileInputRef.current?.click()}
              className={`
                border-2 border-dashed rounded-3xl p-12 text-center cursor-pointer transition-all
                ${file ? 'border-emerald-200 bg-emerald-50/30' : 'border-stone-200 bg-white hover:border-stone-300 hover:bg-stone-50/50'}
              `}
            >
              <input 
                type="file" 
                ref={fileInputRef} 
                onChange={handleFileChange} 
                className="hidden" 
                accept=".pdf,image/*"
              />
              <div className="w-16 h-16 bg-white rounded-2xl shadow-sm border border-stone-100 flex items-center justify-center mx-auto mb-4">
                {isExtracting ? (
                  <Loader2 className="w-8 h-8 animate-spin text-emerald-600" />
                ) : (
                  <Upload className="w-8 h-8 text-stone-400" />
                )}
              </div>
              {file ? (
                <div>
                  <p className="font-semibold text-lg">{file.name}</p>
                  <p className="text-stone-500 text-sm">คลิกเพื่อเปลี่ยนไฟล์</p>
                </div>
              ) : (
                <div>
                  <p className="font-semibold text-lg">อัปโหลดเอกสาร WP</p>
                  <p className="text-stone-500 text-sm">ไฟล์ PDF หรือรูปภาพของใบอนุญาตทำงาน</p>
                </div>
              )}
            </div>
          </section>

          <AnimatePresence mode="wait">
            {error && (
              <motion.div 
                key="error-message"
                initial={{ opacity: 0, height: 0 }}
                animate={{ opacity: 1, height: 'auto' }}
                exit={{ opacity: 0, height: 0 }}
                className="bg-red-50 border border-red-100 rounded-2xl p-4 flex items-start gap-3 text-red-700"
              >
                <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
                <div className="text-sm font-medium">
                  <p>{error}</p>
                  {error.includes('Google Calendar API') && (
                    <p className="mt-2 text-xs opacity-80">
                      Please enable the Google Calendar API in your Google Cloud Console.
                    </p>
                  )}
                </div>
              </motion.div>
            )}

            {extractedData && extractedData.events.length > 0 && !result && (
              <motion.div 
                key="extraction-result"
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                className="bg-white border border-stone-200 rounded-3xl overflow-hidden shadow-sm"
              >
                <div className="p-6 border-b border-stone-100 bg-stone-50/50 flex items-center justify-between">
                  <h3 className="font-bold flex items-center gap-2">
                    <FileText className="w-4 h-4 text-emerald-600" />
                    ข้อมูลที่สกัดได้ (พบ {extractedData.events.length} รายการ)
                  </h3>
                  <button 
                    onClick={() => extractData(file!)}
                    className="text-stone-400 hover:text-stone-900 transition-colors"
                    title="ลองใหม่"
                  >
                    <RefreshCw className="w-4 h-4" />
                  </button>
                </div>
                <div className="p-8 space-y-8">
                  {extractedData.events.map((event, idx) => (
                    <div key={idx} className={`space-y-6 ${idx !== 0 ? 'pt-8 border-t border-stone-100' : ''}`}>
                      <div className="flex items-center gap-2 mb-2">
                        <span className="w-6 h-6 rounded-full bg-stone-100 flex items-center justify-center text-[10px] font-bold text-stone-500">
                          {idx + 1}
                        </span>
                        <h4 className="font-bold text-stone-700">รายการ: {event.stationName}</h4>
                      </div>
                      <div className="grid grid-cols-2 gap-8">
                        <div>
                          <label className="text-[10px] uppercase tracking-widest font-bold text-stone-400 block mb-1">เลขที่ WP</label>
                          <div className="flex items-center gap-2">
                            <p className="font-mono text-lg font-medium">{nextWp ? nextWp : event.wpNumber}</p>
                            {nextWp && (
                              <span className="text-[9px] bg-emerald-100 text-emerald-700 px-1.5 py-0.5 rounded-md font-bold uppercase tracking-wider">รันเลขอัตโนมัติ</span>
                            )}
                          </div>
                        </div>
                        <div>
                          <label className="text-[10px] uppercase tracking-widest font-bold text-stone-400 block mb-1">วันที่</label>
                          <p className="text-lg font-medium">{event.date}</p>
                        </div>
                        <div className="col-span-2">
                          <label className="text-[10px] uppercase tracking-widest font-bold text-stone-400 block mb-1">ชื่อสถานี</label>
                          <p className="text-lg font-medium">{event.stationName}</p>
                        </div>
                        <div className="col-span-2">
                          <label className="text-[10px] uppercase tracking-widest font-bold text-stone-400 block mb-1">หน่วยงานที่ขออนุญาตทำงาน</label>
                          <p className="text-lg font-medium">{event.requestingUnit}</p>
                        </div>
                        <div className="col-span-2">
                          <label className="text-[10px] uppercase tracking-widest font-bold text-stone-400 block mb-1">รายละเอียดงาน</label>
                          <p className="text-lg font-medium">{event.workDescription}</p>
                        </div>
                        <div>
                          <label className="text-[10px] uppercase tracking-widest font-bold text-stone-400 block mb-1">เวลาเริ่ม</label>
                          <p className="text-sm font-medium">{event.startTime}</p>
                        </div>
                        <div>
                          <label className="text-[10px] uppercase tracking-widest font-bold text-stone-400 block mb-1">เวลาสิ้นสุด</label>
                          <p className="text-sm font-medium">{event.endTime}</p>
                        </div>
                      </div>
                    </div>
                  ))}

                  <div className="pt-4 border-t border-stone-100">
                    {!isAuthenticated ? (
                      <button 
                        onClick={() => handleLogin()}
                        className="w-full bg-stone-900 text-white py-4 rounded-2xl font-bold hover:bg-stone-800 transition-all flex items-center justify-center gap-2 shadow-lg shadow-stone-200 active:scale-[0.98]"
                      >
                        <Calendar className="w-5 h-5" />
                        เชื่อมต่อ Google เพื่อดำเนินการต่อ
                      </button>
                    ) : (
                      <button 
                        onClick={handleProcess}
                        disabled={isProcessing}
                        className="w-full bg-emerald-600 text-white py-4 rounded-2xl font-bold hover:bg-emerald-700 transition-all disabled:opacity-50 flex items-center justify-center gap-2 shadow-lg shadow-emerald-100 active:scale-[0.98]"
                      >
                        {isProcessing ? (
                          <>
                            <Loader2 className="w-5 h-5 animate-spin" />
                            กำลังดำเนินการ {extractedData.events.length} รายการ...
                          </>
                        ) : (
                          <>
                            <CheckCircle2 className="w-5 h-5" />
                            ยืนยันและดำเนินการทั้งหมด
                          </>
                        )}
                      </button>
                    )}
                  </div>
                </div>
              </motion.div>
            )}

            {result && (
              <motion.div 
                key="automation-success"
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                className="bg-emerald-600 text-white rounded-3xl p-10 shadow-xl shadow-emerald-200"
              >
                <div className="flex items-center gap-4 mb-8">
                  <div className="w-12 h-12 bg-white/20 rounded-full flex items-center justify-center">
                    <CheckCircle2 className="w-6 h-6" />
                  </div>
                  <div>
                    <h3 className="text-2xl font-bold">ดำเนินการสำเร็จ</h3>
                    <p className="text-emerald-100">เปลี่ยนชื่อไฟล์, อัปโหลด และสร้างกิจกรรมในปฏิทินเรียบร้อยแล้ว</p>
                    {result.sheetError && (
                      <div className="mt-2 text-sm bg-amber-500/20 p-3 rounded-lg border border-amber-400/30 text-amber-100">
                        <p className="font-bold flex items-center gap-2 mb-1">
                          ⚠️ บันทึกลง Google Sheet ไม่สำเร็จ
                        </p>
                        <p className="mb-2 opacity-90">{result.sheetError}</p>
                        
                        {result.sheetError.toLowerCase().includes('permission') && (
                          <div className="mt-2 pt-2 border-t border-amber-400/20 text-xs space-y-1">
                            <p className="font-semibold text-amber-200">วิธีแก้ไข:</p>
                            {isServiceAccount ? (
                              <p>1. แชร์ไฟล์ Google Sheet ให้กับอีเมล Service Account: <code className="bg-black/20 px-1 rounded select-all">{serviceAccountEmail}</code> โดยให้สิทธิ์เป็น "Editor"</p>
                            ) : (
                              <p>1. ตรวจสอบว่าคุณมีสิทธิ์เขียน (Editor) ในไฟล์ Google Sheet นี้</p>
                            )}
                            <p>2. ตรวจสอบว่า GOOGLE_SHEET_ID ในการตั้งค่าถูกต้อง</p>
                            {!isServiceAccount && <p>3. ลองออกจากระบบแล้วเข้าใหม่ และตรวจสอบว่าได้เลือกสิทธิ์ "See, edit, create, and delete all your Google Sheets spreadsheets"</p>}
                          </div>
                        )}
                      </div>
                    )}
                    {result.calendarError && (
                      <div className="mt-2 text-sm bg-amber-500/20 p-3 rounded-lg border border-amber-400/30 text-amber-100">
                        <p className="font-bold flex items-center gap-2 mb-1">
                          ⚠️ สร้างกิจกรรมในปฏิทินไม่สำเร็จ
                        </p>
                        <p className="mb-2 opacity-90">{result.calendarError}</p>
                        
                        {(result.calendarError.toLowerCase().includes('permission') || result.calendarError.toLowerCase().includes('scope')) && (
                          <div className="mt-2 pt-2 border-t border-amber-400/20 text-xs space-y-1">
                            <p className="font-semibold text-amber-200">วิธีแก้ไข:</p>
                            <p>1. ลองออกจากระบบแล้วเข้าใหม่ และตรวจสอบว่าได้เลือกสิทธิ์จัดการปฏิทินทั้งหมด</p>
                            <p>2. ตรวจสอบว่าบัญชีของคุณมีสิทธิ์สร้างกิจกรรมในปฏิทิน</p>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                </div>
                
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <a 
                    href={result.driveLink} 
                    target="_blank" 
                    rel="noopener noreferrer"
                    className="bg-white/10 hover:bg-white/20 border border-white/20 p-6 rounded-2xl transition-all flex flex-col gap-2 group"
                  >
                    <div className="flex items-center justify-between">
                      <Upload className="w-5 h-5 text-emerald-200" />
                      <ExternalLink className="w-4 h-4 opacity-0 group-hover:opacity-100 transition-opacity" />
                    </div>
                    <span className="font-bold">ดูใน Drive</span>
                  </a>
                  <a 
                    href={result.calendarLink} 
                    target="_blank" 
                    rel="noopener noreferrer"
                    className="bg-white/10 hover:bg-white/20 border border-white/20 p-6 rounded-2xl transition-all flex flex-col gap-2 group"
                  >
                    <div className="flex items-center justify-between">
                      <Calendar className="w-5 h-5 text-emerald-200" />
                      <ExternalLink className="w-4 h-4 opacity-0 group-hover:opacity-100 transition-opacity" />
                    </div>
                    <span className="font-bold">ดูในปฏิทิน</span>
                  </a>
                  {result.sheetLink && (
                    <a 
                      href={result.sheetLink} 
                      target="_blank" 
                      rel="noopener noreferrer"
                      className="bg-white/10 hover:bg-white/20 border border-white/20 p-6 rounded-2xl transition-all flex flex-col gap-2 group sm:col-span-2"
                    >
                      <div className="flex items-center justify-between">
                        <FileText className="w-5 h-5 text-emerald-200" />
                        <ExternalLink className="w-4 h-4 opacity-0 group-hover:opacity-100 transition-opacity" />
                      </div>
                      <span className="font-bold">ดูใน Google Sheets</span>
                    </a>
                  )}
                </div>

                <button 
                  onClick={resetState}
                  className="mt-8 w-full py-4 border border-white/30 rounded-2xl font-bold hover:bg-white/10 transition-all"
                >
                  ดำเนินการไฟล์อื่น
                </button>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </main>

      {/* Footer */}
      <footer className="max-w-5xl mx-auto px-6 py-12 border-t border-stone-200 mt-12">
        <div className="flex flex-col gap-8">
          {!isAuthenticated && (
            <div className="bg-amber-50 border border-amber-200 rounded-2xl p-6 text-amber-800 text-sm">
              <h4 className="font-bold mb-2 flex items-center gap-2">
                <AlertCircle className="w-4 h-4" />
                จำเป็นต้องตั้งค่า Google OAuth
              </h4>
              <p className="mb-4">หากคุณเห็นข้อความ "redirect_uri_mismatch" กรุณาเพิ่ม URL นี้ใน Google Cloud Console "Authorized redirect URIs":</p>
              <div className="bg-white border border-amber-200 p-3 rounded-xl font-mono break-all select-all">
                {window.location.origin}/auth/callback
              </div>
            </div>
          )}
          
          <div className="flex flex-col md:flex-row items-center justify-between gap-6">
            <p className="text-stone-400 text-sm">
              © {new Date().getFullYear()} ระบบจัดการ WP อัตโนมัติ. พัฒนาด้วย Gemini AI.
            </p>
            <div className="flex items-center gap-6 text-stone-400 text-sm font-medium">
              <span className="flex items-center gap-1.5">
                <div className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
                ระบบพร้อมใช้งาน
              </span>
            </div>
          </div>
        </div>
      </footer>
    </div>
  );
}
