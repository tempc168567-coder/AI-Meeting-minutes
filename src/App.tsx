/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef, useCallback } from 'react';
import { Mic, Square, Sparkles, Mail, Clock, FileText, ChevronRight, X, LogIn, LogOut, Download } from 'lucide-react';
import { initAuth, googleSignIn, getAccessToken, logout } from './lib/auth';
import type { User } from 'firebase/auth';
import { Document, Packer, Paragraph, HeadingLevel } from 'docx';
import { saveAs } from 'file-saver';

// --- Types ---
interface Meeting {
  id: string;
  startTime: number;
  endTime: number | null;
  transcript: string;
  summary: string;
  title?: string;
}

// --- Icons / Components ---
const Button = ({ children, onClick, disabled, variant = 'default', className = '', type = 'button', title }: any) => {
  const base = "inline-flex items-center justify-center gap-2 px-6 py-3 text-sm font-bold transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500 disabled:pointer-events-none disabled:opacity-50 rounded-lg";
  const variants: any = {
    default: "bg-indigo-600 text-white hover:bg-indigo-700 shadow-lg shadow-indigo-100",
    destructive: "bg-red-600 text-white hover:bg-red-700 shadow-lg shadow-red-100",
    secondary: "bg-white text-slate-600 border border-slate-200 hover:bg-slate-50 font-semibold shadow-sm",
    ghost: "text-slate-600 hover:bg-slate-100 font-semibold",
  };
  return (
    <button type={type} onClick={onClick} disabled={disabled} className={`${base} ${variants[variant]} ${className}`} title={title}>
      {children}
    </button>
  );
};

export default function App() {
  const [meetings, setMeetings] = useState<Meeting[]>([]);
  const [selectedMeetingId, setSelectedMeetingId] = useState<string | null>(null);
  
  const [isRecording, setIsRecording] = useState(false);
  const [isSummarizing, setIsSummarizing] = useState(false);
  const [interimResult, setInterimResult] = useState('');
  
  const [isEmailSectionOpen, setIsEmailSectionOpen] = useState(false);
  const [emailRecipient, setEmailRecipient] = useState('');
  const [emailMessage, setEmailMessage] = useState('');
  const [isSendingEmail, setIsSendingEmail] = useState(false);
  const [emailSuccess, setEmailSuccess] = useState('');
  const [emailError, setEmailError] = useState('');

  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [isAuthChecking, setIsAuthChecking] = useState(true);
  const [authError, setAuthError] = useState('');

  const handleSignIn = async () => {
    try {
      setAuthError('');
      const result = await googleSignIn();
      if (result) {
        setCurrentUser(result.user);
      }
    } catch (err: any) {
      console.error('Login error:', err);
      if (err.code === 'auth/popup-closed-by-user') {
        setAuthError('登入視窗已關閉。');
      } else {
        setAuthError('登入失敗，您的瀏覽器可能阻擋了彈出視窗(Popup)或不支援此環境。強烈建議點選右上角的「Open App in new tab (箭頭圖示)」在新分頁開啟應用程式以完成登入！');
      }
    }
  };

  const recognitionRef = useRef<any>(null);
  const shouldRecordRef = useRef<boolean>(false);
  const selectedMeetingIdRef = useRef<string | null>(null);

  useEffect(() => {
    selectedMeetingIdRef.current = selectedMeetingId;
    if (selectedMeetingId) {
       // Also reset email form state when switching meetings
       setIsEmailSectionOpen(false);
       setEmailRecipient('');
       setEmailMessage('');
       setEmailSuccess('');
       setEmailError('');
    }
  }, [selectedMeetingId]);

  // Setup auth
  useEffect(() => {
    const unsubscribe = initAuth(
      (user) => {
        setCurrentUser(user);
        setIsAuthChecking(false);
      },
      () => {
        setCurrentUser(null);
        setIsAuthChecking(false);
      }
    );
    return () => unsubscribe();
  }, []);

  // Initialize Speech Recognition
  useEffect(() => {
    // @ts-ignore
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (SpeechRecognition) {
      const recognition = new SpeechRecognition();
      recognition.continuous = true;
      recognition.interimResults = true;
      recognition.lang = 'zh-TW';

      recognition.onresult = (event: any) => {
        let finalStr = '';
        let interimStr = '';
        for (let i = event.resultIndex; i < event.results.length; i++) {
          if (event.results[i].isFinal) {
            finalStr += event.results[i][0].transcript + ' ';
          } else {
            interimStr += event.results[i][0].transcript;
          }
        }
        
        if (finalStr) {
          setMeetings(prev => prev.map(m => {
            if (m.id === selectedMeetingIdRef.current) {
              return { ...m, transcript: (m.transcript || '') + finalStr };
            }
            return m;
          }));
        }
        setInterimResult(interimStr);
      };

      recognition.onerror = (event: any) => {
        console.error('Speech recognition error', event.error);
        if (event.error === 'not-allowed') {
          alert('請允許麥克風權限以使用聽寫功能。');
          setIsRecording(false);
          shouldRecordRef.current = false;
        }
      };

      recognition.onend = () => {
        if (shouldRecordRef.current) {
          try {
            recognition.start();
          } catch(e) {
            console.error(e);
          }
        }
      };

      recognitionRef.current = recognition;
    } else {
      console.warn('SpeechRecognition API is not supported in this browser.');
    }
  }, []);

  // Load from local storage
  useEffect(() => {
    const stored = localStorage.getItem('meeting_history');
    if (stored) {
      try {
        const parsed = JSON.parse(stored);
        setMeetings(parsed);
        if (parsed.length > 0) {
          setSelectedMeetingId(parsed[0].id);
        }
      } catch (e) {
        console.error("Failed to parse history", e);
      }
    }
  }, []);

  // Save to local storage on changes
  useEffect(() => {
    localStorage.setItem('meeting_history', JSON.stringify(meetings));
  }, [meetings]);

  const startNewMeeting = useCallback(() => {
    const newMeeting: Meeting = {
      id: crypto.randomUUID(),
      startTime: Date.now(),
      endTime: null,
      transcript: '',
      summary: '',
      title: `會議紀錄 ${new Date().toLocaleDateString('zh-TW')} ${new Date().toLocaleTimeString('zh-TW', { hour: '2-digit', minute: '2-digit' })}`
    };
    setMeetings(prev => [newMeeting, ...prev]);
    setSelectedMeetingId(newMeeting.id);
    selectedMeetingIdRef.current = newMeeting.id; // Immediate update
    
    // Start recording automatically
    setIsRecording(true);
    shouldRecordRef.current = true;
    setInterimResult('');
    try {
      recognitionRef.current?.start();
    } catch (e) {
      console.log('Recognition already started');
    }
  }, []);

  const stopRecording = useCallback(() => {
    setIsRecording(false);
    shouldRecordRef.current = false;
    setInterimResult('');
    recognitionRef.current?.stop();
    
    setMeetings(prev => prev.map(m => {
      if (m.id === selectedMeetingIdRef.current) {
        return { ...m, endTime: m.endTime || Date.now() };
      }
      return m;
    }));
  }, []);

  const generateSummary = async () => {
    const meeting = meetings.find(m => m.id === selectedMeetingId);
    if (!meeting || !meeting.transcript.trim()) return;

    setIsSummarizing(true);
    try {
      const response = await fetch('/api/summarize', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ transcript: meeting.transcript })
      });
      
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || 'Failed to summarize');

      setMeetings(prev => prev.map(m => {
        if (m.id === selectedMeetingId) {
          return { ...m, summary: data.summary };
        }
        return m;
      }));
    } catch (err: any) {
      alert("生成摘要失敗: " + err.message);
    } finally {
      setIsSummarizing(false);
    }
  };

  const handleSummaryChange = (newSummary: string) => {
    setMeetings(prev => prev.map(m => {
      if (m.id === selectedMeetingId) {
        return { ...m, summary: newSummary };
      }
      return m;
    }));
  };

  const handleTranscriptChange = (newTranscript: string) => {
    setMeetings(prev => prev.map(m => {
      if (m.id === selectedMeetingId) {
        return { ...m, transcript: newTranscript };
      }
      return m;
    }));
  };

  const activeMeeting = meetings.find(m => m.id === selectedMeetingId);

  const hasValidEmails = React.useMemo(() => {
    if (!emailRecipient.trim()) return false;
    const emails = emailRecipient.split(/[,;]/).map(e => e.trim()).filter(e => e);
    const emailRegex = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
    return emails.length > 0 && emails.every(e => emailRegex.test(e));
  }, [emailRecipient]);

  const handleSendEmail = async (e: React.FormEvent) => {
    e.preventDefault();
    if (isSendingEmail) return;
    
    if (!currentUser) {
      setEmailError('請先登入 Google 帳號授權寄信。');
      return;
    }
    
    if (!hasValidEmails) {
      setEmailError('請輸入有效的收件者電子郵件。');
      return;
    }
    
    if (emailMessage.trim().length === 0) {
      setEmailError('郵件內容不能為空。');
      return;
    }

    setIsSendingEmail(true);
    setEmailSuccess('');
    setEmailError('');

    try {
      const token = await getAccessToken();
      if (!token) throw new Error('無法取得授權 Token，請重新登入。');

      const subject = `[會議紀錄] ${activeMeeting?.title || '會議'}`;
      
      const emailLines = [];
      emailLines.push(`To: ${emailRecipient}`);
      emailLines.push(`Subject: =?utf-8?B?${btoa(unescape(encodeURIComponent(subject)))}?=`);
      emailLines.push('Content-Type: text/plain; charset=utf-8');
      emailLines.push('');
      emailLines.push(emailMessage);
      
      const emailRaw = btoa(unescape(encodeURIComponent(emailLines.join('\r\n'))))
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=+$/, '');

      const response = await fetch('https://gmail.googleapis.com/gmail/v1/users/me/messages/send', {
        method: 'POST',
        headers: { 
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json' 
        },
        body: JSON.stringify({ raw: emailRaw })
      });
      
      const data = await response.json();
      if (!response.ok) throw new Error(data.error?.message || 'Failed to send email');

      setEmailSuccess('信件已成功透過您的 Gmail 寄出！');
      setEmailRecipient('');
    } catch (err: any) {
      setEmailError(err.message || '無法寄送，請稍後再試。');
    } finally {
      setIsSendingEmail(false);
    }
  };

  const handleToggleEmailBlock = () => {
     if (!isEmailSectionOpen) {
        // Pre-fill email message from summary/transcript
        const prefill = activeMeeting?.summary ? `=== 摘要重點 ===\n${activeMeeting.summary}` : (activeMeeting?.transcript ? `=== 會議聽寫內容 ===\n${activeMeeting.transcript.substring(0, 500)}...` : '');
        setEmailMessage(prefill);
     }
     setIsEmailSectionOpen(!isEmailSectionOpen);
  };

  const handleExportWord = async () => {
    if (!activeMeeting) return;

    const title = activeMeeting.title || '會議紀錄';
    const children = [];

    children.push(
      new Paragraph({
        text: title,
        heading: HeadingLevel.HEADING_1,
        spacing: { after: 300 },
      })
    );

    children.push(
      new Paragraph({
        text: `時間：${formatDate(activeMeeting.startTime)}`,
        spacing: { after: 300 },
      })
    );

    if (activeMeeting.summary) {
      children.push(
        new Paragraph({
          text: '摘要重點',
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 200, after: 100 },
        })
      );
      activeMeeting.summary.split('\n').forEach(line => {
        children.push(new Paragraph({ text: line, spacing: { after: 100 } }));
      });
    }

    if (activeMeeting.transcript) {
      children.push(
        new Paragraph({
          text: '會議聽寫內容',
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 300, after: 100 },
        })
      );
      activeMeeting.transcript.split('\n').forEach(line => {
        children.push(new Paragraph({ text: line, spacing: { after: 100 } }));
      });
    }

    const doc = new Document({
      sections: [{ children }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `${title}.docx`);
  };

  const formatDate = (ts: number) => {
    return new Intl.DateTimeFormat('zh-TW', {
      year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit'
    }).format(new Date(ts));
  };

  return (
    <div className="flex flex-col h-screen bg-slate-50 text-slate-900 font-sans overflow-hidden relative">
      
      {/* Header / Status Bar */}
      <header className="bg-white border-b border-slate-200 px-6 py-3 flex justify-between items-center shrink-0 z-10">
        <div className="flex items-center gap-3">
          {isRecording ? <div className="w-3 h-3 bg-red-500 rounded-full animate-pulse"></div> : <Mic className="w-4 h-4 text-slate-400" />}
          <h1 className="text-lg font-bold tracking-tight text-slate-800">
            AI 會議助手 
            {isRecording && <span className="text-slate-400 font-normal ml-2">| 即時轉錄中...</span>}
          </h1>
        </div>
        <div className="flex items-center gap-4 text-sm font-medium text-slate-500">
          <span>{activeMeeting ? formatDate(activeMeeting.startTime) : '歡迎使用'}</span>
          <div className="w-px h-4 bg-slate-300 mx-2"></div>
          {isAuthChecking ? (
            <span className="text-slate-400">驗證中...</span>
          ) : currentUser ? (
            <div className="flex items-center gap-2">
              <span className="text-slate-700 font-semibold">{currentUser.displayName}</span>
              <button onClick={logout} className="text-xs text-slate-400 hover:text-slate-600 transition-colors" title="登出">登出</button>
            </div>
          ) : (
            <button
               onClick={handleSignIn}
               className="flex items-center gap-1.5 text-indigo-600 hover:text-indigo-800 font-bold transition-colors"
            >
              <LogIn className="w-4 h-4" /> 登入以寄送紀錄
            </button>
          )}
        </div>
      </header>

      {activeMeeting ? (
        <>
          {/* Main Content Area */}
          <main className="flex-1 p-6 flex gap-6 overflow-hidden min-h-0">
            {/* Dictation */}
            <section className="flex-[2] bg-white rounded-xl border border-slate-200 shadow-sm flex flex-col">
              <div className="p-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center shrink-0">
                <span className="text-xs font-bold uppercase tracking-wider text-slate-400">即時聽寫內容 {(!isRecording && activeMeeting.transcript) ? ' (可手動修改)' : ''}</span>
                <span className="text-xs text-indigo-600 font-semibold">自動儲存至雲端</span>
              </div>
              <div className="flex-1 p-6 flex flex-col overflow-hidden leading-relaxed text-slate-700 font-sans text-sm">
                {isRecording ? (
                  <div className="flex-1 overflow-y-auto whitespace-pre-wrap">
                    {activeMeeting.transcript || <span className="text-slate-400 font-medium italic">（語音辨識中，尚無內容...）</span>}
                    {interimResult && <span className="text-indigo-400 italic"> {interimResult}</span>}
                  </div>
                ) : (
                  <textarea
                    className="flex-1 w-full h-full resize-none border-none focus:ring-0 bg-transparent outline-none placeholder:text-slate-400 p-0 m-0 whitespace-pre-wrap"
                    value={activeMeeting.transcript || ''}
                    onChange={(e) => handleTranscriptChange(e.target.value)}
                    placeholder="尚無語音紀錄，可在此手動輸入內容..."
                  />
                )}
              </div>
            </section>

            {/* Right Column */}
            <div className="flex-[1.2] flex flex-col gap-6 overflow-hidden min-h-0">
              {/* Summary */}
              <section className="flex-1 bg-white rounded-xl border border-slate-200 shadow-sm flex flex-col min-h-0">
                <div className="p-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center shrink-0">
                  <span className="text-xs font-bold uppercase tracking-wider text-slate-400">摘要重點編輯</span>
                  <Sparkles className="w-4 h-4 text-slate-400" />
                </div>
                <div className="p-5 flex-1 flex flex-col overflow-hidden">
                  <textarea
                    className="w-full h-full resize-none border-none focus:ring-0 text-sm leading-relaxed text-slate-600 bg-transparent outline-none placeholder:text-slate-300"
                    value={activeMeeting.summary || ''}
                    onChange={(e) => handleSummaryChange(e.target.value)}
                    placeholder="請點擊下方「生成摘要」按鈕，或在此直接輸入摘要內容..."
                  />
                </div>
              </section>
            </div>
          </main>

          {/* Action Bar */}
          <nav className="h-24 bg-white border-y border-slate-200 px-8 flex items-center justify-between shrink-0">
            <div className="flex gap-3">
              {isRecording ? (
                <Button variant="destructive" onClick={stopRecording}>
                  <div className="w-2 h-2 bg-white rounded-full"></div> 結束會議並儲存
                </Button>
              ) : (
                <Button onClick={() => {
                   if (activeMeeting.endTime) {
                     startNewMeeting();
                   } else {
                     setIsRecording(true);
                     shouldRecordRef.current = true;
                     recognitionRef.current?.start();
                   }
                }}>
                  <Mic className="w-4 h-4" /> {activeMeeting.endTime ? '新會議' : '繼續錄音'}
                </Button>
              )}

              <Button 
                variant="default" 
                onClick={generateSummary} 
                disabled={!activeMeeting.transcript || isRecording || isSummarizing}
              >
                <Sparkles className="w-4 h-4" />
                {isSummarizing ? '生成中...' : '生成摘要'}
              </Button>
            </div>
            <div className="flex gap-3">
              <Button 
                variant="secondary"
                onClick={handleExportWord}
                disabled={!activeMeeting.transcript && !activeMeeting.summary}
              >
                <Download className="w-4 h-4 text-slate-600" /> 匯出 Word
              </Button>
              <Button 
                variant="secondary"
                onClick={handleToggleEmailBlock}
                className={isEmailSectionOpen ? 'ring-2 ring-indigo-500 bg-indigo-50 border-indigo-200 text-indigo-700' : ''}
              >
                <Mail className={`w-4 h-4 ${isEmailSectionOpen ? 'text-indigo-600' : 'text-slate-600'}`} /> {isEmailSectionOpen ? '關閉 Email' : '寄送 Email'}
              </Button>
            </div>
          </nav>
        </>
      ) : (
        /* Empty State */
        <main className="flex-1 flex flex-col items-center justify-center p-6 bg-slate-50 min-h-0">
          <div className="w-16 h-16 bg-indigo-50 text-indigo-600 rounded-2xl flex items-center justify-center mb-6 shadow-sm ring-1 ring-indigo-200">
            <Mic className="w-8 h-8" />
          </div>
          <h2 className="text-2xl font-bold tracking-tight text-slate-800 mb-2">準備開始新會議</h2>
          <p className="text-slate-500 max-w-sm mb-8 leading-relaxed text-center text-sm font-medium">
            使用即時語音轉文字功能，並透過 AI 自動生成專業會議摘要。
          </p>
          <Button onClick={startNewMeeting} className="px-8 py-3 text-base shadow-lg shadow-indigo-100">
            開始第一場會議
          </Button>
        </main>
      )}

      {/* History List Bottom Section */}
      <section className="h-64 bg-slate-50 p-6 shrink-0 z-10 border-t border-slate-200">
        <div className="max-w-full h-full bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden flex flex-col">
          <div className="px-6 py-3 border-b border-slate-100 flex justify-between items-center bg-white shrink-0">
            <h2 className="text-sm font-bold text-slate-700">歷史會議記錄</h2>
            <div className="flex items-center gap-4">
              <span className="text-xs text-slate-400 font-bold tracking-wider">共 {meetings.length} 筆記錄</span>
            </div>
          </div>
          <div className="overflow-y-auto flex-1 bg-white">
            {meetings.length === 0 ? (
              <div className="text-center py-12 text-slate-400 text-sm font-semibold">
                尚無歷史會議紀錄。
              </div>
            ) : (
              <table className="w-full text-left border-collapse">
                <thead className="bg-slate-50 text-[10px] uppercase font-bold text-slate-400 sticky top-0 z-10">
                  <tr>
                    <th className="px-6 py-2">會議名稱</th>
                    <th className="px-6 py-2 w-48">時間</th>
                    <th className="px-6 py-2">摘要預覽</th>
                    <th className="px-6 py-2 w-24 text-right">操作</th>
                  </tr>
                </thead>
                <tbody className="text-sm divide-y divide-slate-100">
                  {meetings.map((m) => (
                    <tr 
                      key={m.id} 
                      onClick={() => {
                        if (isRecording) stopRecording();
                        setSelectedMeetingId(m.id);
                        selectedMeetingIdRef.current = m.id;
                      }}
                      className={`cursor-pointer group transition-colors ${
                        selectedMeetingId === m.id ? 'bg-indigo-50/50' : 'hover:bg-indigo-50'
                      }`}
                    >
                      <td className="px-6 py-3 font-semibold text-slate-700">
                        <div className="flex items-center gap-2">
                          {m.title}
                          {selectedMeetingId === m.id && <div className="w-1.5 h-1.5 rounded-full bg-indigo-500"></div>}
                        </div>
                      </td>
                      <td className="px-6 py-3 text-slate-500 text-xs font-medium">
                        {formatDate(m.startTime)}
                      </td>
                      <td className="px-6 py-3 text-slate-400 truncate max-w-[300px]">
                        {m.summary ? m.summary.substring(0, 50) + '...' : m.transcript ? m.transcript.substring(0, 50) + '...' : '尚無內容'}
                      </td>
                      <td className="px-6 py-3 text-right">
                        <span className="text-indigo-600 font-bold opacity-0 group-hover:opacity-100 text-xs uppercase tracking-wider transition-opacity">
                          {selectedMeetingId === m.id ? '目前查看' : '載入'}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      </section>

      {/* Email Modal */}
      {isEmailSectionOpen && (
        <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center z-50">
          <div className="bg-white w-[500px] rounded-2xl shadow-2xl overflow-hidden border border-slate-300 animate-in fade-in zoom-in-95 duration-200">
            <div className="p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50">
              <h3 className="text-lg font-bold flex items-center gap-2 text-slate-800"><Mail className="w-5 h-5 text-indigo-600"/> 寄送會議紀錄</h3>
              <button onClick={handleToggleEmailBlock} className="text-slate-400 hover:text-slate-600 transition-colors" title="關閉">
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="p-6 flex flex-col gap-4">
              {emailSuccess && (
                <div className="p-3 bg-green-50 text-green-700 text-sm font-medium rounded-lg border border-green-200 animate-in fade-in transition-all">
                  {emailSuccess}
                </div>
              )}
              {emailError && (
                <div className="p-3 bg-red-50 text-red-700 text-sm font-medium rounded-lg border border-red-200 animate-in fade-in flex justify-between items-start transition-all">
                  <span>{emailError}</span>
                  <button onClick={() => setEmailError('')} className="text-red-500 hover:text-red-700"><X className="w-4 h-4" /></button>
                </div>
              )}
              <div>
                <label className="block text-xs font-bold text-slate-500 mb-1.5">收件者 (多筆請用逗號隔開)</label>
                <input
                  type="text"
                  value={emailRecipient}
                  onChange={(e) => setEmailRecipient(e.target.value)}
                  placeholder="example@company.com"
                  className={`w-full p-2.5 bg-slate-50 border ${emailRecipient && !hasValidEmails ? 'border-red-300 focus:border-red-500 focus:ring-red-200' : 'border-slate-200 focus:border-indigo-500 focus:ring-indigo-200'} rounded-lg text-sm text-slate-700 focus:ring-2 outline-none transition-all`}
                />
                {emailRecipient && !hasValidEmails && (
                   <p className="text-red-500 text-xs mt-1.5 font-medium">請輸入有效的電子郵件格式</p>
                )}
              </div>
              {!currentUser && (
                 <div className="mt-4 p-3 bg-amber-50 text-amber-700 text-sm font-medium rounded-lg border border-amber-200">
                    您需要登入 Google 帳號才能寄送電子郵件。若登入失敗，請嘗試開啟用新分頁開啟此應用程式。
                 </div>
              )}
              {authError && (
                 <div className="mt-2 p-3 bg-red-50 text-red-700 text-sm font-medium rounded-lg border border-red-200">
                    {authError}
                 </div>
              )}
              <div className="flex flex-col min-h-[160px]">
                <label className="block text-xs font-bold text-slate-500 mb-1.5">摘要 / 內容</label>
                <textarea
                  value={emailMessage}
                  onChange={(e) => setEmailMessage(e.target.value)}
                  className="w-full flex-1 p-3 bg-slate-50 border border-slate-200 rounded-lg text-sm text-slate-700 focus:ring-2 focus:ring-indigo-200 focus:border-indigo-500 outline-none transition-all resize-none"
                  placeholder="在此輸入要寄送的內容..."
                />
              </div>
            </div>
            <div className="p-6 border-t border-slate-100 bg-slate-50 flex justify-end gap-3 rounded-b-xl">
              <Button variant="ghost" onClick={handleToggleEmailBlock} disabled={isSendingEmail} className="px-4 py-2 font-semibold">
                 取消
              </Button>
              <Button 
                 variant="default" 
                 onClick={handleSendEmail} 
                 disabled={isSendingEmail || !currentUser}
                 className="px-6 py-2 shadow-md flex items-center gap-2"
              >
                {isSendingEmail ? (
                  <>
                    <div className="w-4 h-4 rounded-full border-2 border-white/30 border-t-white animate-spin"></div>
                    寄送中...
                  </>
                ) : (
                  <>
                    <Mail className="w-4 h-4" /> 確認寄出
                  </>
                )}
              </Button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}

