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
  FileCheck
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

// Initialize Gemini
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

interface ExtractedData {
  wpNumber: string;
  stationName: string;
  date: string;
  isoDate: string;
  calendarTitle: string;
  isStaffed: boolean;
}

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState<boolean | null>(null);
  const [file, setFile] = useState<File | null>(null);
  const [isExtracting, setIsExtracting] = useState(false);
  const [extractedData, setExtractedData] = useState<ExtractedData | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [result, setResult] = useState<{ driveLink: string; calendarLink: string } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    checkAuthStatus();
    
    const handleMessage = (event: MessageEvent) => {
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        setIsAuthenticated(true);
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

  const checkAuthStatus = async () => {
    try {
      const res = await fetch('/api/auth/status');
      const data = await res.json();
      setIsAuthenticated(data.isAuthenticated);
    } catch (err) {
      console.error('Auth check failed', err);
      setIsAuthenticated(false);
    }
  };

  const handleLogin = async () => {
    try {
      const res = await fetch('/api/auth/url');
      const { url } = await res.json();
      window.open(url, 'google_oauth', 'width=600,height=700');
    } catch (err) {
      setError('Failed to get auth URL');
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
      const reader = new FileReader();
      const fileDataPromise = new Promise<string>((resolve) => {
        reader.onload = () => resolve((reader.result as string).split(',')[1]);
        reader.readAsDataURL(file);
      });
      const base64Data = await fileDataPromise;

      const response = await ai.models.generateContent({
        model: "gemini-2.5-flash",
        contents: [
          {
            parts: [
              {
                inlineData: {
                  mimeType: file.type,
                  data: base64Data
                }
              },
              {
                text: `Extract information from this Work Permit (WP) document for a Thai electrical substation.
                Find:
                1. WP Number (e.g., 013-69)
                2. Station Name (e.g., สถานีไฟฟ้ากระทุ่มแบน 6)
                3. Date of work (Thai format, e.g., 13 ก.พ. 69)
                4. ISO Date (YYYY-MM-DD)
                5. Is it staffed (จัดพนักงาน) or unstaffed (ไม่จัดพนักงาน)? Look for keywords like "จัดพนักงาน" or "ไม่จัดพนักงาน". Default to "จัดพนักงาน" if unsure.
                
                Generate a Calendar Title pattern: "[Station Name] ([Staffed Status]) WP ผจฟ.1 No.[WP Number] บำรุงรักษาระบบ SCPS ประจำปี (ผปค.กสฟ.ก3)"
                
                Return JSON only.`
              }
            ]
          }
        ],
        config: {
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              wpNumber: { type: Type.STRING },
              stationName: { type: Type.STRING },
              date: { type: Type.STRING },
              isoDate: { type: Type.STRING },
              isStaffed: { type: Type.BOOLEAN },
              calendarTitle: { type: Type.STRING }
            },
            required: ["wpNumber", "stationName", "date", "isoDate", "isStaffed", "calendarTitle"]
          }
        }
      });

      const data = JSON.parse(response.text || '{}');
      setExtractedData(data);
    } catch (err: any) {
      console.error('Extraction failed', err);
      setError('Failed to extract data from file. Please try again or check the file format.');
    } finally {
      setIsExtracting(false);
    }
  };

  const handleProcess = async () => {
    if (!file || !extractedData) return;
    setIsProcessing(true);
    setError(null);

    const formData = new FormData();
    formData.append('file', file);
    formData.append('wpNumber', extractedData.wpNumber);
    formData.append('stationName', extractedData.stationName);
    formData.append('date', extractedData.date);
    formData.append('isoDate', extractedData.isoDate);
    formData.append('calendarTitle', extractedData.calendarTitle);

    try {
      const res = await fetch('/api/process', {
        method: 'POST',
        body: formData
      });
      const data = await res.json();
      if (data.success) {
        setResult({
          driveLink: data.driveFile.webViewLink,
          calendarLink: data.calendarEvent.htmlLink
        });
      } else {
        setError(data.error || 'Processing failed');
      }
    } catch (err) {
      setError('Connection error. Please try again.');
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

  return (
    <div className="min-h-screen bg-stone-50 text-stone-900 font-sans selection:bg-emerald-100">
      {/* Header */}
      <header className="border-b border-stone-200 bg-white/80 backdrop-blur-md sticky top-0 z-10">
        <div className="max-w-5xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-emerald-600 rounded-lg flex items-center justify-center text-white">
              <FileCheck className="w-5 h-5" />
            </div>
            <h1 className="font-semibold text-lg tracking-tight">WP Automation</h1>
          </div>
          {isAuthenticated && (
            <button 
              onClick={handleLogout}
              className="flex items-center gap-2 text-sm font-medium text-stone-500 hover:text-stone-900 transition-colors"
            >
              <LogOut className="w-4 h-4" />
              Sign Out
            </button>
          )}
        </div>
      </header>

      <main className="max-w-3xl mx-auto px-6 py-12">
        {!isAuthenticated ? (
          <motion.div 
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="bg-white border border-stone-200 rounded-3xl p-12 text-center shadow-sm"
          >
            <div className="w-20 h-20 bg-emerald-50 rounded-full flex items-center justify-center mx-auto mb-6">
              <Calendar className="w-10 h-10 text-emerald-600" />
            </div>
            <h2 className="text-3xl font-bold mb-4 tracking-tight">Connect to Google</h2>
            <p className="text-stone-500 mb-8 max-w-md mx-auto leading-relaxed">
              We need access to your Google Drive and Calendar to automate your WP workflow.
            </p>
            <button 
              onClick={handleLogin}
              className="bg-stone-900 text-white px-8 py-4 rounded-2xl font-semibold hover:bg-stone-800 transition-all shadow-lg shadow-stone-200 active:scale-95"
            >
              Sign in with Google
            </button>
          </motion.div>
        ) : (
          <div className="space-y-8">
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
                    <p className="text-stone-500 text-sm">Click to change file</p>
                  </div>
                ) : (
                  <div>
                    <p className="font-semibold text-lg">Upload WP Document</p>
                    <p className="text-stone-500 text-sm">PDF or Image of the Work Permit</p>
                  </div>
                )}
              </div>
            </section>

            <AnimatePresence mode="wait">
              {error && (
                <motion.div 
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: 'auto' }}
                  exit={{ opacity: 0, height: 0 }}
                  className="bg-red-50 border border-red-100 rounded-2xl p-4 flex items-start gap-3 text-red-700"
                >
                  <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
                  <p className="text-sm font-medium">{error}</p>
                </motion.div>
              )}

              {extractedData && !result && (
                <motion.div 
                  initial={{ opacity: 0, y: 20 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="bg-white border border-stone-200 rounded-3xl overflow-hidden shadow-sm"
                >
                  <div className="p-6 border-b border-stone-100 bg-stone-50/50 flex items-center justify-between">
                    <h3 className="font-bold flex items-center gap-2">
                      <FileText className="w-4 h-4 text-emerald-600" />
                      Extracted Information
                    </h3>
                    <button 
                      onClick={() => extractData(file!)}
                      className="text-stone-400 hover:text-stone-900 transition-colors"
                      title="Retry Extraction"
                    >
                      <RefreshCw className="w-4 h-4" />
                    </button>
                  </div>
                  <div className="p-8 space-y-6">
                    <div className="grid grid-cols-2 gap-8">
                      <div>
                        <label className="text-[10px] uppercase tracking-widest font-bold text-stone-400 block mb-1">WP Number</label>
                        <p className="font-mono text-lg font-medium">{extractedData.wpNumber}</p>
                      </div>
                      <div>
                        <label className="text-[10px] uppercase tracking-widest font-bold text-stone-400 block mb-1">Date</label>
                        <p className="text-lg font-medium">{extractedData.date}</p>
                      </div>
                      <div className="col-span-2">
                        <label className="text-[10px] uppercase tracking-widest font-bold text-stone-400 block mb-1">Station Name</label>
                        <p className="text-lg font-medium">{extractedData.stationName}</p>
                      </div>
                    </div>

                    <div className="space-y-4 pt-4 border-t border-stone-100">
                      <div>
                        <label className="text-[10px] uppercase tracking-widest font-bold text-stone-400 block mb-1">New File Name Pattern</label>
                        <div className="bg-stone-50 p-3 rounded-xl text-sm font-mono text-stone-600 break-all">
                          WP ผจฟ.1 No.{extractedData.wpNumber} กสฟ.(ก3) เข้า {extractedData.stationName} ({extractedData.date})
                        </div>
                      </div>
                      <div>
                        <label className="text-[10px] uppercase tracking-widest font-bold text-stone-400 block mb-1">Calendar Title Pattern</label>
                        <div className="bg-stone-50 p-3 rounded-xl text-sm font-mono text-stone-600 break-all">
                          {extractedData.calendarTitle}
                        </div>
                      </div>
                    </div>

                    <button 
                      onClick={handleProcess}
                      disabled={isProcessing}
                      className="w-full bg-emerald-600 text-white py-4 rounded-2xl font-bold hover:bg-emerald-700 transition-all disabled:opacity-50 flex items-center justify-center gap-2 shadow-lg shadow-emerald-100 active:scale-[0.98]"
                    >
                      {isProcessing ? (
                        <>
                          <Loader2 className="w-5 h-5 animate-spin" />
                          Processing...
                        </>
                      ) : (
                        <>
                          <CheckCircle2 className="w-5 h-5" />
                          Confirm & Automate
                        </>
                      )}
                    </button>
                  </div>
                </motion.div>
              )}

              {result && (
                <motion.div 
                  initial={{ opacity: 0, scale: 0.95 }}
                  animate={{ opacity: 1, scale: 1 }}
                  className="bg-emerald-600 text-white rounded-3xl p-10 shadow-xl shadow-emerald-200"
                >
                  <div className="flex items-center gap-4 mb-8">
                    <div className="w-12 h-12 bg-white/20 rounded-full flex items-center justify-center">
                      <CheckCircle2 className="w-6 h-6" />
                    </div>
                    <div>
                      <h3 className="text-2xl font-bold">Automation Complete</h3>
                      <p className="text-emerald-100">File renamed, uploaded, and calendar event created.</p>
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
                      <span className="font-bold">View in Drive</span>
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
                      <span className="font-bold">View in Calendar</span>
                    </a>
                  </div>

                  <button 
                    onClick={resetState}
                    className="mt-8 w-full py-4 border border-white/30 rounded-2xl font-bold hover:bg-white/10 transition-all"
                  >
                    Process Another File
                  </button>
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="max-w-5xl mx-auto px-6 py-12 border-t border-stone-200 mt-12">
        <div className="flex flex-col md:flex-row items-center justify-between gap-6">
          <p className="text-stone-400 text-sm">
            © {new Date().getFullYear()} WP Automation Tool. Built with Gemini AI.
          </p>
          <div className="flex items-center gap-6 text-stone-400 text-sm font-medium">
            <span className="flex items-center gap-1.5">
              <div className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
              System Active
            </span>
          </div>
        </div>
      </footer>
    </div>
  );
}
