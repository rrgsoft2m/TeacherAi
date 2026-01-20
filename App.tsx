
import React, { useState, useEffect } from 'react';
import { View, LessonData, User, ResultTab, AuthMode, HistoryItem, HistoryType } from './types';
import { Icons } from './constants';
import { generateLesson, generateStandaloneImage, generateLogicPuzzle } from './geminiService';

declare const PptxGenJS: any;
declare const docx: any;
declare const saveAs: any;

const Logo = ({ size = 'md' }: { size?: 'sm' | 'md' | 'lg' }) => (
  <div className={`flex flex-col items-center text-center ${size === 'lg' ? 'space-y-4 md:space-y-6' : 'space-y-2'}`}>
    <div className={`${size === 'lg' ? 'w-20 h-20 md:w-24 md:h-24' : 'w-10 h-10 md:w-14 md:h-14'} flex items-center justify-center filter drop-shadow-xl`}>
      <img src="https://buxedu.uz/static/images/logo.png" alt="Logo" className="w-full h-full object-contain" />
    </div>
    <div className="max-w-[200px] md:max-w-xs">
      <h1 className={`${size === 'lg' ? 'text-sm md:text-lg font-extrabold' : 'text-[8px] md:text-[10px] font-extrabold'} text-[#1e3a8a] uppercase tracking-tighter leading-tight`}>
        MAKTABGACHA VA MAKTAB<br/>TA'LIMI VAZIRLIGI
      </h1>
    </div>
  </div>
);

const App: React.FC = () => {
  const [view, setView] = useState<View>(View.Loading);
  const [authMode, setAuthMode] = useState<AuthMode>(AuthMode.Login);
  const [authError, setAuthError] = useState<string | null>(null);
  const [user, setUser] = useState<User | null>(null);
  const [history, setHistory] = useState<HistoryItem[]>([]);
  const [currentLesson, setCurrentLesson] = useState<LessonData | null>(null);
  const [activeTab, setActiveTab] = useState<ResultTab>(ResultTab.Presentation);
  const [formData, setFormData] = useState({ 
    subject: '', 
    grade: '', 
    topic: '', 
    goal: 'Yangi mavzu', 
    language: "O'zbekcha" 
  });
  
  const [imagePrompt, setImagePrompt] = useState('');
  const [generatedImage, setGeneratedImage] = useState<string | null>(null);
  const [puzzleTopic, setPuzzleTopic] = useState('');
  const [generatedPuzzle, setGeneratedPuzzle] = useState<{puzzle: string, solution: string} | null>(null);
  const [isGeneratingExtra, setIsGeneratingExtra] = useState(false);

  const [userAnswers, setUserAnswers] = useState<number[]>(new Array(10).fill(-1));
  const [showTestResults, setShowTestResults] = useState(false);

  useEffect(() => {
    const savedSession = localStorage.getItem('teacher_ai_session');
    if (savedSession) {
      const parsedUser = JSON.parse(savedSession);
      setUser(parsedUser);
      loadHistory(parsedUser.email);
      setView(View.Landing);
    } else {
      setView(View.Auth);
    }
  }, []);

  const loadHistory = (email: string) => {
    const saved = localStorage.getItem(`history_${email}`);
    if (saved) setHistory(JSON.parse(saved));
    else setHistory([]);
  };

  const saveToHistory = (type: HistoryType, title: string, data: any) => {
    if (!user) return;
    const newItem: HistoryItem = {
      id: Math.random().toString(36).substr(2, 9),
      type,
      date: new Date().toLocaleString('uz-UZ'),
      title,
      data
    };
    const updatedHistory = [newItem, ...history];
    setHistory(updatedHistory);
    localStorage.setItem(`history_${user.email}`, JSON.stringify(updatedHistory));
  };

  const handleAuth = (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    setAuthError(null);
    const fd = new FormData(e.currentTarget);
    const email = fd.get('email') as string;
    const password = fd.get('password') as string;
    const name = fd.get('name') as string;

    const users = JSON.parse(localStorage.getItem('teacher_ai_users') || '[]');

    if (authMode === AuthMode.Register) {
      if (!email || !password || !name) {
        setAuthError("Iltimos, barcha maydonlarni to'ldiring!");
        return;
      }
      if (users.find((u: any) => u.email === email)) {
        setAuthError("Bu email manzili allaqachon mavjud!");
        return;
      }
      const newUser = { email, password, name, role: "PREMIUM O'QITUVCHI", isPremium: true };
      users.push(newUser);
      localStorage.setItem('teacher_ai_users', JSON.stringify(users));
      localStorage.setItem('teacher_ai_session', JSON.stringify(newUser));
      setUser(newUser);
      setHistory([]);
      setView(View.Landing);
    } else {
      const found = users.find((u: any) => u.email === email && u.password === password);
      if (found) {
        localStorage.setItem('teacher_ai_session', JSON.stringify(found));
        setUser(found);
        loadHistory(email);
        setView(View.Landing);
      } else {
        setAuthError("Email yoki maxfiy parol noto'g'ri!");
      }
    }
  };

  const exportPPTX = async () => {
    if (!currentLesson) return;
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';

    let slide = pptx.addSlide();
    slide.addText(currentLesson.topic, { x: 1, y: 1.5, w: '80%', fontSize: 40, bold: true, color: '1e3a8a', align: 'center' });
    slide.addText(`${currentLesson.subject} | ${currentLesson.grade}`, { x: 1, y: 2.5, w: '80%', fontSize: 20, color: '64748b', align: 'center' });
    slide.addText("MAKTABGACHA VA MAKTAB TA'LIMI VAZIRLIGI", { x: 1, y: 4.5, w: '80%', fontSize: 14, color: 'cbd5e1', align: 'center', italic: true });

    currentLesson.slides.forEach(s => {
      let sld = pptx.addSlide();
      sld.addText(s.title, { x: 0.5, y: 0.5, w: '90%', fontSize: 28, bold: true, color: '1e3a8a' });
      sld.addText(s.content, { x: 0.5, y: 1.5, w: '50%', fontSize: 16, color: '475569' });
      if (s.imageUrl) {
        sld.addImage({ data: s.imageUrl, x: 5.5, y: 1.0, w: 4.2, h: 3.2 });
      }
    });

    pptx.writeFile({ fileName: `${currentLesson.topic}_Taqdimot.pptx` });
  };

  const exportDOCX = async (mode: 'TEST' | 'QA') => {
    if (!currentLesson) return;
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = docx;

    const children = [
      new Paragraph({ 
        text: currentLesson.topic.toUpperCase(), 
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 }
      }),
      new Paragraph({ 
        text: mode === 'TEST' ? "INTERAKTIV TESTLAR" : "SAVOL-JAVOBLAR", 
        heading: HeadingLevel.HEADING_2,
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 }
      }),
      new Paragraph({ 
        children: [
          new TextRun({ text: `Fan: ${currentLesson.subject} | Sinf: ${currentLesson.grade} | Til: ${currentLesson.language}`, bold: true, size: 24 })
        ],
        spacing: { after: 400 }
      }),
    ];

    if (mode === 'TEST') {
      currentLesson.tests.forEach((t, i) => {
        children.push(new Paragraph({ 
          children: [new TextRun({ text: `${i + 1}. ${t.question}`, bold: true, size: 26 })],
          spacing: { before: 200, after: 100 }
        }));
        t.options.forEach((opt, oi) => {
          children.push(new Paragraph({ text: `   ${String.fromCharCode(65+oi)}) ${opt}`, size: 24 }));
        });
        children.push(new Paragraph({ 
          children: [
            new TextRun({ text: `Javob: ${String.fromCharCode(65+t.correctIndex)}) ${t.options[t.correctIndex]}`, bold: true, color: '22c55e', size: 24 })
          ],
          spacing: { before: 100 }
        }));
      });
    } else {
      currentLesson.qa.forEach((q, i) => {
        children.push(new Paragraph({ 
          children: [new TextRun({ text: `${i + 1}. Savol: ${q.question}`, bold: true, size: 26 })],
          spacing: { before: 200, after: 100 }
        }));
        children.push(new Paragraph({ text: `Javob: ${q.answer}`, size: 24 }));
      });
    }

    const doc = new Document({ sections: [{ children }] });
    const blob = await Packer.toBlob(doc);
    saveAs(blob, `${currentLesson.topic}_${mode}.docx`);
  };

  const handleGenerateImage = async () => {
    if (!imagePrompt) return;
    setIsGeneratingExtra(true);
    try {
      const url = await generateStandaloneImage(imagePrompt);
      setGeneratedImage(url);
      saveToHistory('IMAGE', `AI Rasm: ${imagePrompt}`, url);
    } catch (e) { alert("Xatolik yuz berdi!"); }
    finally { setIsGeneratingExtra(false); }
  };

  const handleGeneratePuzzle = async () => {
    if (!puzzleTopic) return;
    setIsGeneratingExtra(true);
    try {
      const res = await generateLogicPuzzle(puzzleTopic);
      setGeneratedPuzzle(res);
      saveToHistory('PUZZLE', `Jumboq: ${puzzleTopic}`, res);
    } catch (e) { alert("Xatolik yuz berdi!"); }
    finally { setIsGeneratingExtra(false); }
  };

  if (view === View.Loading) return (
    <div className="min-h-screen flex flex-col items-center justify-center bg-white p-4">
      <Logo size="lg" />
      <div className="mt-12 brain-loader"></div>
      <p className="mt-6 text-slate-400 font-bold uppercase tracking-widest text-xs">Ilova yuklanmoqda...</p>
    </div>
  );

  if (view === View.Auth) return (
    <div className="min-h-screen flex items-center justify-center p-4 bg-slate-50">
      <div className="main-card w-full max-w-md p-6 md:p-10 border border-white shadow-2xl flex flex-col items-center">
        <Logo size="sm" />
        <h2 className="text-2xl font-black text-slate-800 mt-8 mb-6">{authMode === AuthMode.Login ? 'Xush kelibsiz!' : "Ro'yxatdan o'ting"}</h2>
        <form onSubmit={handleAuth} className="w-full space-y-4">
          {authMode === AuthMode.Register && <input name="name" type="text" placeholder="Ism Familiya" className="w-full glass-input rounded-xl py-4 px-6 font-bold" required />}
          <input name="email" type="email" placeholder="Email manzili" className="w-full glass-input rounded-xl py-4 px-6 font-bold" required />
          <input name="password" type="password" placeholder="Maxfiy parol" className="w-full glass-input rounded-xl py-4 px-6 font-bold" required />
          {authError && <p className="text-red-500 text-xs font-bold text-center animate-shake">{authError}</p>}
          <button type="submit" className="w-full btn-blue py-4 rounded-xl font-black uppercase shadow-xl transition-all active:scale-95">
            {authMode === AuthMode.Login ? 'Tizimga kirish' : "Ro'yxatdan o'tish"}
          </button>
        </form>
        <button onClick={() => { setAuthMode(authMode === AuthMode.Login ? AuthMode.Register : AuthMode.Login); setAuthError(null); }} className="mt-6 text-[10px] font-bold text-blue-500 uppercase tracking-widest hover:underline decoration-2 underline-offset-4">
          {authMode === AuthMode.Login ? 'YANGI HISOB YARATISH →' : '← TIZIMGA KIRISH'}
        </button>
      </div>
    </div>
  );

  return (
    <div className="min-h-screen flex flex-col bg-[#f9fbff]">
      <header className="fixed top-0 left-0 right-0 z-[100] px-2 md:px-6 py-4">
        <div className="max-w-7xl mx-auto flex items-center justify-between bg-white/80 backdrop-blur-xl rounded-full px-4 md:px-8 py-3 shadow-lg border border-white/50">
          <div className="flex items-center space-x-3 cursor-pointer group" onClick={() => setView(View.Landing)}>
            <div className="w-10 h-10 rounded-full bg-blue-600 text-white flex items-center justify-center font-black text-lg shadow-xl group-hover:scale-110 transition-all">{user?.name[0]}</div>
            <div className="hidden sm:flex flex-col">
              <span className="font-bold text-slate-800 text-sm">{user?.name}</span>
              <span className="text-[10px] text-blue-500 font-bold uppercase tracking-widest">PREMIUM USTOZ</span>
            </div>
          </div>
          <div className="flex space-x-1 md:space-x-3">
            <button onClick={() => setView(View.CreateForm)} className={`px-4 md:px-6 py-2 rounded-full text-xs md:text-sm font-bold transition-all ${view === View.CreateForm ? 'bg-blue-600 text-white shadow-lg' : 'hover:bg-slate-50 text-slate-500'}`}>
              <span className="md:hidden">✨</span><span className="hidden md:inline">Yaratish</span>
            </button>
            <button onClick={() => setView(View.Dashboard)} className={`px-4 md:px-6 py-2 rounded-full text-xs md:text-sm font-bold transition-all ${view === View.Dashboard ? 'bg-blue-600 text-white shadow-lg' : 'hover:bg-slate-50 text-slate-500'}`}>
               <span className="md:hidden">📁</span><span className="hidden md:inline">Tarix ({history.length})</span>
            </button>
            <button onClick={() => { localStorage.removeItem('teacher_ai_session'); setView(View.Auth); }} className="w-10 h-10 flex items-center justify-center text-slate-300 hover:text-red-500 transition-colors text-xl">🚪</button>
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-7xl mx-auto w-full px-4 pt-32 pb-20">
        {view === View.Landing && (
          <div className="min-h-[60vh] flex flex-col items-center justify-center text-center space-y-12 animate-in fade-in zoom-in duration-700">
            <div className="space-y-6">
              <h1 className="text-4xl md:text-7xl font-black text-slate-900 leading-tight tracking-tighter">Darsingizni <span className="text-blue-600">AI</span> bilan<br/>modernizatsiya qiling!</h1>
              <p className="text-slate-500 text-lg md:text-xl max-w-3xl mx-auto italic font-medium">"Maktab o'qituvchilari uchun maxsus sun'iy intellektga asoslangan dars ishlanmalari portali"</p>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-6 w-full max-w-4xl">
              <div onClick={() => setView(View.CreateForm)} className="main-card p-10 cursor-pointer hover:-translate-y-2 transition-all hover:shadow-2xl border border-slate-50 group">
                 <div className="text-6xl mb-6 group-hover:scale-110 transition-transform">⚡</div>
                 <h3 className="text-xl font-black text-slate-800">Tezkor dars paketi</h3>
                 <p className="mt-2 text-xs text-slate-400 font-bold uppercase tracking-widest">Slayd + Test + O'yin</p>
              </div>
              <div onClick={() => setView(View.AiImage)} className="main-card p-10 cursor-pointer hover:-translate-y-2 transition-all hover:shadow-2xl border border-slate-50 group">
                 <div className="text-6xl mb-6 group-hover:scale-110 transition-transform">🪄</div>
                 <h3 className="text-xl font-black text-slate-800">AI Rasm Generator</h3>
                 <p className="mt-2 text-xs text-slate-400 font-bold uppercase tracking-widest">Dars uchun 3D vizuallar</p>
              </div>
              <div onClick={() => setView(View.LogicGame)} className="main-card p-10 cursor-pointer hover:-translate-y-2 transition-all hover:shadow-2xl border border-slate-50 group">
                 <div className="text-6xl mb-6 group-hover:scale-110 transition-transform">🧠</div>
                 <h3 className="text-xl font-black text-slate-800">Mantiqiy Jumboqlar</h3>
                 <p className="mt-2 text-xs text-slate-400 font-bold uppercase tracking-widest">Kreativ savol-javoblar</p>
              </div>
            </div>
          </div>
        )}

        {view === View.Dashboard && (
          <div className="space-y-8 animate-in slide-in-from-bottom-6 duration-500">
             <div className="flex items-center space-x-4">
                <div className="w-1.5 h-10 bg-blue-600 rounded-full"></div>
                <h2 className="text-3xl font-black text-slate-900">Mening barcha ishlarim</h2>
             </div>
             <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                {history.length === 0 ? (
                  <div className="col-span-full py-20 text-center text-slate-300 font-black uppercase tracking-widest">Hozircha tarix bo'sh.</div>
                ) : history.map(item => (
                  <div key={item.id} onClick={() => {
                    if (item.type === 'LESSON') { setCurrentLesson(item.data); setView(View.Results); }
                    else if (item.type === 'IMAGE') { setGeneratedImage(item.data); setView(View.AiImage); }
                    else { setGeneratedPuzzle(item.data); setView(View.LogicGame); }
                  }} className="main-card p-6 cursor-pointer group border border-slate-100 hover:border-blue-400 transition-all hover:shadow-2xl">
                     <div className="flex justify-between items-start mb-6">
                        <span className={`px-4 py-1.5 rounded-full text-[10px] font-black uppercase tracking-widest shadow-sm ${
                          item.type === 'LESSON' ? 'bg-blue-600 text-white' : item.type === 'IMAGE' ? 'bg-purple-600 text-white' : 'bg-orange-500 text-white'
                        }`}>{item.type}</span>
                        <span className="text-[10px] text-slate-400 font-bold">{item.date.split(',')[0]}</span>
                     </div>
                     <h3 className="text-lg font-black text-slate-800 line-clamp-2 mb-4 group-hover:text-blue-600 transition-colors leading-tight">{item.title}</h3>
                     <div className="flex items-center text-xs text-slate-400 font-bold uppercase tracking-widest">
                       <span>Ko'rish →</span>
                     </div>
                  </div>
                ))}
             </div>
          </div>
        )}

        {view === View.AiImage && (
          <div className="max-w-4xl mx-auto space-y-8 animate-in fade-in">
             <button onClick={() => setView(View.Landing)} className="flex items-center space-x-2 font-black text-slate-400 hover:text-blue-600 transition-colors uppercase text-xs tracking-widest"><span>←</span><span>Bosh sahifaga</span></button>
             <div className="main-card p-8 md:p-16 space-y-12">
                <h2 className="text-4xl font-black text-slate-900">AI Rasm Generator</h2>
                <div className="flex flex-col sm:flex-row gap-4">
                   <input value={imagePrompt} onChange={e => setImagePrompt(e.target.value)} placeholder="Tavsif bering..." className="flex-1 glass-input rounded-2xl py-5 px-8 font-bold text-lg" />
                   <button onClick={handleGenerateImage} disabled={isGeneratingExtra} className="btn-blue px-12 py-5 rounded-2xl font-black uppercase shadow-xl disabled:opacity-50">{isGeneratingExtra ? '...' : 'Yaratish'}</button>
                </div>
                {generatedImage && <div className="rounded-[48px] overflow-hidden shadow-2xl border-[12px] border-white animate-in zoom-in duration-1000"><img src={generatedImage} className="w-full h-auto" /></div>}
             </div>
          </div>
        )}

        {view === View.LogicGame && (
          <div className="max-w-4xl mx-auto space-y-8 animate-in fade-in">
             <button onClick={() => setView(View.Landing)} className="flex items-center space-x-2 font-black text-slate-400 hover:text-blue-600 transition-colors uppercase text-xs tracking-widest"><span>←</span><span>Bosh sahifaga</span></button>
             <div className="main-card p-8 md:p-16 space-y-12">
                <h2 className="text-4xl font-black text-slate-900">Mantiqiy Jumboqlar</h2>
                <div className="flex flex-col sm:flex-row gap-4">
                   <input value={puzzleTopic} onChange={e => setPuzzleTopic(e.target.value)} placeholder="Mavzu nomi..." className="flex-1 glass-input rounded-2xl py-5 px-8 font-bold text-lg" />
                   <button onClick={handleGeneratePuzzle} disabled={isGeneratingExtra} className="btn-blue px-12 py-5 rounded-2xl font-black uppercase shadow-xl disabled:opacity-50">{isGeneratingExtra ? '...' : 'Savol olish'}</button>
                </div>
                {generatedPuzzle && (
                  <div className="bg-blue-50/50 p-8 md:p-12 rounded-[48px] border-4 border-white space-y-8 animate-in fade-in shadow-inner">
                     <p className="text-2xl md:text-3xl font-medium text-slate-800 leading-relaxed italic">"{generatedPuzzle.puzzle}"</p>
                     <details className="bg-white p-6 rounded-3xl border-2 border-blue-100 cursor-pointer shadow-sm group">
                        <summary className="font-black text-blue-600 uppercase tracking-widest outline-none">Javobni ko'rish</summary>
                        <p className="mt-4 text-xl font-bold text-slate-700 italic border-l-4 border-blue-500 pl-4">{generatedPuzzle.solution}</p>
                     </details>
                  </div>
                )}
             </div>
          </div>
        )}

        {view === View.CreateForm && (
          <div className="max-w-4xl mx-auto main-card p-8 md:p-16 space-y-10 shadow-2xl border border-white animate-in zoom-in-95">
             <h2 className="text-4xl font-black text-slate-900">Yangi dars rejasi</h2>
             <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-2">
                   <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-4">O'quv fani</label>
                   <input value={formData.subject} onChange={e => setFormData({...formData, subject: e.target.value})} placeholder="Biologiya..." className="w-full glass-input rounded-2xl py-4 px-6 font-bold" />
                </div>
                <div className="space-y-2">
                   <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-4">Sinf</label>
                   <input value={formData.grade} onChange={e => setFormData({...formData, grade: e.target.value})} placeholder="7-sinf..." className="w-full glass-input rounded-2xl py-4 px-6 font-bold" />
                </div>
                <div className="md:col-span-2 space-y-2">
                   <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-4">Dars mavzusi</label>
                   <input value={formData.topic} onChange={e => setFormData({...formData, topic: e.target.value})} placeholder="Mavzuni to'liq yozing..." className="w-full glass-input rounded-[28px] py-6 px-10 font-bold text-xl" />
                </div>
                {/* Language Select */}
                <div className="space-y-2">
                   <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-4">Til tanlash</label>
                   <select value={formData.language} onChange={e => setFormData({...formData, language: e.target.value})} className="w-full glass-input rounded-2xl py-4 px-6 font-bold cursor-pointer">
                      <option value="O'zbekcha">🇺🇿 O'zbekcha</option>
                      <option value="Ruscha">🇷🇺 Ruscha</option>
                      <option value="Inglizcha">🇺🇸 Inglizcha</option>
                   </select>
                </div>
                {/* Goal Select */}
                <div className="space-y-2">
                   <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-4">Dars maqsadi</label>
                   <select value={formData.goal} onChange={e => setFormData({...formData, goal: e.target.value})} className="w-full glass-input rounded-2xl py-4 px-6 font-bold cursor-pointer">
                      <option value="Yangi mavzu">📚 Yangi mavzu</option>
                      <option value="Takrorlash">🔄 Takrorlash</option>
                      <option value="Imtihon">📝 Imtihon</option>
                   </select>
                </div>
             </div>
             <button onClick={async () => {
                if (!formData.subject || !formData.topic) return alert("Barcha maydonlarni to'ldiring!");
                setView(View.Generating);
                try {
                  const lesson = await generateLesson(formData);
                  setCurrentLesson(lesson);
                  saveToHistory('LESSON', lesson.topic, lesson);
                  setView(View.Results);
                  setUserAnswers(new Array(10).fill(-1));
                  setShowTestResults(false);
                } catch (e) { alert("Xatolik!"); setView(View.CreateForm); }
             }} className="w-full btn-blue py-6 rounded-[32px] font-black uppercase text-xl shadow-2xl">YARATISH 🚀</button>
          </div>
        )}

        {view === View.Generating && (
          <div className="flex flex-col items-center justify-center py-32 space-y-12">
             <div className="brain-loader border-[8px]"></div>
             <h2 className="text-4xl font-black text-slate-900">Materiallar tayyorlanmoqda...</h2>
          </div>
        )}

        {view === View.Results && currentLesson && (
          <div className="space-y-12 animate-in fade-in duration-700">
             {/* Sticky Navigation */}
             <div className="sticky top-[84px] z-[90] py-4 bg-[#f9fbff]/80 backdrop-blur-md border-b border-slate-100">
                <div className="max-w-7xl mx-auto flex justify-center px-2">
                   <div className="bg-white p-2 rounded-full shadow-2xl border-4 border-white flex space-x-1 md:space-x-2 overflow-x-auto no-scrollbar max-w-full">
                      {Object.values(ResultTab).map(tab => (
                        <button key={tab} onClick={() => setActiveTab(tab)} className={`px-4 md:px-10 py-3 rounded-full text-[9px] md:text-xs font-black uppercase whitespace-nowrap transition-all ${activeTab === tab ? 'bg-blue-600 text-white shadow-xl scale-105' : 'text-slate-400 hover:bg-slate-50'}`}>{tab}</button>
                      ))}
                   </div>
                </div>
             </div>

             <div className="main-card p-6 md:p-16 border border-white shadow-2xl min-h-[600px]">
                {activeTab === ResultTab.Presentation && (
                  <div className="space-y-20">
                     <div className="flex justify-between items-center flex-wrap gap-4 border-b pb-10">
                        <h2 className="text-3xl font-black">📊 Taqdimot Slaydlari</h2>
                        <button onClick={exportPPTX} className="bg-orange-500 text-white px-8 py-3.5 rounded-full font-black text-xs uppercase shadow-xl hover:bg-orange-600 transition-all flex items-center space-x-2">
                          <span>📊</span><span>PowerPoint (.pptx) yuklash</span>
                        </button>
                     </div>
                     <div className="space-y-32">
                        {currentLesson.slides.map((s, i) => (
                          <div key={i} className={`flex flex-col lg:flex-row gap-12 items-center ${i % 2 === 0 ? '' : 'lg:flex-row-reverse'} animate-in fade-in`}>
                             <div className="flex-1 space-y-6">
                                <span className="w-14 h-14 bg-blue-900 text-white rounded-[24px] flex items-center justify-center font-black text-2xl shadow-xl">{i+1}</span>
                                <h3 className="text-3xl md:text-4xl font-black text-slate-900 leading-tight">{s.title}</h3>
                                <p className="text-slate-500 text-xl italic border-l-8 border-slate-100 pl-8 bg-slate-50/30 py-4 rounded-r-3xl">{s.content}</p>
                             </div>
                             {s.imageUrl && (
                               <div className="w-full lg:w-[500px] aspect-[4/3] rounded-[56px] overflow-hidden shadow-2xl border-[16px] border-white transition-all hover:scale-[1.03]">
                                 <img src={s.imageUrl} className="w-full h-full object-cover" alt={s.title} />
                               </div>
                             )}
                          </div>
                        ))}
                     </div>
                  </div>
                )}

                {activeTab === ResultTab.Test && (
                  <div className="space-y-12">
                     <div className="flex justify-between items-center flex-wrap gap-4 border-b pb-10">
                        <h2 className="text-3xl font-black">📝 Interaktiv Testlar</h2>
                        <button onClick={() => exportDOCX('TEST')} className="bg-blue-600 text-white px-8 py-3.5 rounded-full font-black text-xs uppercase shadow-xl hover:bg-blue-700 transition-all flex items-center space-x-2">
                          <span>📝</span><span>Word (.docx) yuklash</span>
                        </button>
                     </div>
                     <div className="space-y-8">
                        {currentLesson.tests.map((q, i) => (
                          <div key={i} className={`p-8 md:p-12 rounded-[48px] border-4 transition-all ${showTestResults ? (userAnswers[i] === q.correctIndex ? 'bg-green-50 border-green-200' : 'bg-red-50 border-red-200') : 'bg-white border-slate-50 shadow-xl'}`}>
                             <h4 className="text-2xl font-black mb-10 leading-snug">{i+1}. {q.question}</h4>
                             <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                {q.options.map((opt, oi) => (
                                  <button key={oi} disabled={showTestResults} onClick={() => { const na = [...userAnswers]; na[i] = oi; setUserAnswers(na); }} className={`p-6 rounded-[32px] border-2 font-bold flex items-center space-x-4 transition-all ${userAnswers[i] === oi ? 'bg-blue-600 text-white border-blue-700' : 'bg-slate-50 border-transparent hover:bg-slate-100 text-slate-600'}`}>
                                     <span className="w-12 h-12 rounded-2xl flex items-center justify-center font-black text-lg bg-white/20">{String.fromCharCode(65+oi)}</span>
                                     <span className="text-lg flex-1 text-left">{opt}</span>
                                  </button>
                                ))}
                             </div>
                          </div>
                        ))}
                     </div>
                     {!showTestResults && <button onClick={() => { if(userAnswers.includes(-1)) alert("Hamma savollarga javob bering!"); else setShowTestResults(true); }} className="w-full btn-blue py-6 rounded-[32px] font-black uppercase text-xl shadow-2xl">Tekshirish 🎯</button>}
                  </div>
                )}

                {activeTab === ResultTab.QA && (
                  <div className="space-y-12 animate-in fade-in">
                     <div className="flex justify-between items-center flex-wrap gap-4 border-b pb-10">
                        <h2 className="text-3xl font-black">🗣️ Savol-Javoblar</h2>
                        <button onClick={() => exportDOCX('QA')} className="bg-indigo-600 text-white px-8 py-3.5 rounded-full font-black text-xs uppercase shadow-xl hover:bg-indigo-700 transition-all flex items-center space-x-2">
                          <span>🗣️</span><span>Word (.docx) yuklash</span>
                        </button>
                     </div>
                     <div className="space-y-6">
                        {currentLesson.qa.map((item, i) => (
                          <details key={i} className="group bg-white rounded-[32px] border-2 border-slate-50 shadow-sm overflow-hidden hover:shadow-xl transition-all">
                             <summary className="p-8 cursor-pointer font-black text-xl text-slate-800 list-none flex justify-between items-center">
                                <span className="leading-snug">{i+1}. {item.question}</span>
                                <span className="w-10 h-10 rounded-full bg-slate-50 text-slate-300 flex items-center justify-center group-open:rotate-180 transition-transform shadow-inner">▼</span>
                             </summary>
                             <div className="p-10 pt-0 text-slate-500 text-xl md:text-2xl italic font-medium border-t-2 border-slate-50 bg-slate-50/20 leading-relaxed">
                                "{item.answer}"
                             </div>
                          </details>
                        ))}
                     </div>
                  </div>
                )}

                {activeTab === ResultTab.Interactive && (
                  <div className="space-y-24 animate-in fade-in">
                     <div className="bg-white rounded-[60px] p-8 md:p-16 border-[6px] border-slate-50 shadow-2xl">
                        <h3 className="text-3xl md:text-4xl font-black text-slate-900 mb-16 tracking-tight">🧩 Mavzu Krossvordi</h3>
                        <div className="overflow-x-auto no-scrollbar">
                           <table className="w-full text-left border-collapse min-w-[500px]">
                              <thead>
                                 <tr className="border-b-2 border-slate-100 text-xs uppercase font-black text-blue-600 tracking-[0.5em] opacity-60">
                                    <th className="pb-8 px-4">Ta'rif</th>
                                    <th className="pb-8 px-4 text-right">Kalit so'z</th>
                                 </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100">
                                 {currentLesson.interactive.crossword.map((item, i) => (
                                   <tr key={i} className="group hover:bg-slate-50 transition-colors">
                                      <td className="py-10 px-4 font-bold text-xl md:text-2xl text-slate-800 leading-snug">{item.definition}</td>
                                      <td className="py-10 px-4 text-right font-black text-blue-600 uppercase tracking-[0.5em] text-3xl md:text-4xl group-hover:scale-105 transition-transform origin-right">{item.answer}</td>
                                   </tr>
                                 ))}
                              </tbody>
                           </table>
                        </div>
                     </div>
                     <div className="grid grid-cols-1 md:grid-cols-2 gap-12">
                        <div className="bg-amber-50 p-12 rounded-[60px] border-4 border-amber-100 space-y-10 shadow-inner group hover:shadow-2xl transition-all">
                           <div className="w-20 h-20 bg-amber-500 text-white rounded-[28px] flex items-center justify-center text-4xl shadow-2xl">💡</div>
                           <h3 className="text-3xl font-black text-amber-900">Mantiqiy Jumboq</h3>
                           <p className="text-2xl md:text-3xl font-medium italic text-amber-900/70 leading-relaxed italic border-l-[12px] border-amber-200 pl-8 py-2">
                              "{currentLesson.interactive.puzzle}"
                           </p>
                        </div>
                        <div className="bg-indigo-50 p-12 rounded-[60px] border-4 border-indigo-100 space-y-10 shadow-inner group hover:shadow-2xl transition-all">
                           <div className="w-20 h-20 bg-indigo-600 text-white rounded-[28px] flex items-center justify-center text-4xl shadow-2xl">🎮</div>
                           <h3 className="text-3xl font-black text-indigo-900 tracking-tight">Mini-O'yin:<br/>{currentLesson.interactive.game.name}</h3>
                           <div className="space-y-6">
                              {currentLesson.interactive.game.rules.map((r, i) => (
                                <div key={i} className="flex items-start space-x-4">
                                   <span className="w-10 h-10 bg-indigo-600 text-white rounded-xl flex items-center justify-center font-black shrink-0">{i+1}</span>
                                   <p className="text-xl text-slate-700 font-medium italic pt-1 leading-relaxed">"{r}"</p>
                                </div>
                              ))}
                           </div>
                        </div>
                     </div>
                  </div>
                )}
             </div>
          </div>
        )}
      </main>

      <footer className="w-full max-w-7xl mx-auto px-4 py-20 text-center border-t border-slate-100 mt-20 space-y-10 bg-white/50 backdrop-blur-sm">
        <Logo size="sm" />
        <p className="text-[11px] font-black text-slate-400 uppercase tracking-[0.5em] opacity-60 italic">Kreativ Ta'lim - Milliy Kelajak Poydevori</p>
        <p className="text-[10px] font-black text-slate-500 uppercase tracking-widest leading-loose">
          © 2026 Sirdaryo Viloyati Maktabgacha va Maktab Ta'limi Boshqarmasi <br/>
          <span className="text-slate-300">TEXNIK KO'MAK:</span> <a href="https://t.me/rrgfcoder" target="_blank" className="text-blue-500 underline decoration-blue-100 decoration-4 font-black">RRGSOFT</a>
        </p>
      </footer>
    </div>
  );
};

export default App;
