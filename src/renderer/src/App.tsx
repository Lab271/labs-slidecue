import { useState, useEffect } from 'react';
import { QRCodeSVG } from 'qrcode.react';
import './types/electron.d.ts';
import appIcon from './assets/icon.png';

type AppState = 'checking' | 'no-powerpoint' | 'idle' | 'importing' | 'loaded' | 'presenting';

interface PresentationInfo {
  filePath: string;
  fileName: string;
  thumbnails: string[];
  totalSlides: number;
  visibleSlides: number[];
  hiddenSlides: number[];
}

interface ConnectionInfo {
  url: string;
  pin: string;
}

interface ImportProgress {
  step: number;
  total: number;
  message: string;
}

export default function App() {
  const [state, setState] = useState<AppState>('checking');
  const [presentation, setPresentation] = useState<PresentationInfo | null>(
    null
  );
  const [connection, setConnection] = useState<ConnectionInfo | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [importProgress, setImportProgress] = useState<ImportProgress | null>(null);

  useEffect(() => {
    window.electronAPI.checkPowerPoint().then((installed) => {
      setState(installed ? 'idle' : 'no-powerpoint');
    });

    // Listen for import progress
    const cleanup = window.electronAPI.onImportProgress((data) => {
      setImportProgress(data);
      if (data.step < data.total) {
        setState('importing');
      }
    });

    return cleanup;
  }, []);

  const handleImport = async () => {
    try {
      setError(null);
      const result = await window.electronAPI.importPresentation();
      if (!result) {
        // User cancelled file dialog
        return;
      }
      setPresentation(result);
      setState('loaded');
      setImportProgress(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to import');
      setState('idle');
      setImportProgress(null);
    }
  };

  const handlePresent = async () => {
    try {
      setError(null);
      const result = await window.electronAPI.startPresentation();
      setConnection(result);
      setState('presenting');
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to start');
    }
  };

  const handleStop = async () => {
    try {
      await window.electronAPI.stopPresentation();
      setConnection(null);
      setState('loaded');
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to stop');
    }
  };

  if (state === 'checking') {
    return (
      <div className="flex items-center justify-center h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900">
        <div className="text-center">
          <img src={appIcon} alt="SlideCue" className="w-20 h-20 rounded-2xl mx-auto mb-6 animate-pulse shadow-lg" />
          <div className="w-8 h-8 border-3 border-blue-500/30 border-t-blue-500 rounded-full animate-spin mx-auto mb-4"></div>
          <p className="text-lg text-slate-400">Checking PowerPoint...</p>
        </div>
      </div>
    );
  }

  if (state === 'no-powerpoint') {
    return (
      <div className="flex flex-col items-center justify-center h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900 gap-6 p-8">
        <div className="w-20 h-20 bg-red-500/20 rounded-2xl flex items-center justify-center text-4xl">
          ⚠️
        </div>
        <h1 className="text-2xl font-bold text-red-400">PowerPoint Not Found</h1>
        <p className="text-slate-400 text-center max-w-md">
          Microsoft PowerPoint is required to use SlideCue. Please install
          PowerPoint and restart the app.
        </p>
        <div className="mt-4 px-6 py-3 bg-slate-800/50 rounded-xl border border-slate-700/50">
          <p className="text-slate-500 text-sm">Looking for: Microsoft PowerPoint.app</p>
        </div>
      </div>
    );
  }

  if (state === 'importing' && importProgress) {
    const percentage = Math.round((importProgress.step / importProgress.total) * 100);
    const steps = [
      { label: 'Starting', icon: '🚀' },
      { label: 'Opening', icon: '📂' },
      { label: 'Analyzing', icon: '🔍' },
      { label: 'Converting', icon: '⚙️' },
      { label: 'Finishing', icon: '✨' },
    ];
    const currentStepIndex = Math.min(Math.floor(importProgress.step), steps.length - 1);
    
    return (
      <div className="flex flex-col items-center justify-center h-screen bg-gradient-to-br from-gray-900 via-gray-800 to-gray-900 gap-8 p-8">
        {/* Animated logo/icon */}
        <div className="relative">
          <div className="text-7xl animate-bounce">{steps[currentStepIndex].icon}</div>
          <div className="absolute -inset-4 bg-blue-500/20 rounded-full blur-xl animate-pulse"></div>
        </div>
        
        {/* Title */}
        <div className="text-center">
          <h1 className="text-4xl font-bold text-white mb-2">Preparing Your Slides</h1>
          <p className="text-blue-400 text-lg font-medium">{importProgress.message}</p>
        </div>
        
        {/* Progress bar */}
        <div className="w-96 max-w-full">
          <div className="bg-gray-700/50 rounded-full h-3 overflow-hidden backdrop-blur-sm border border-gray-600/30">
            <div 
              className="bg-gradient-to-r from-blue-500 via-blue-400 to-cyan-400 h-full transition-all duration-500 ease-out relative"
              style={{ width: `${percentage}%` }}
            >
              <div className="absolute inset-0 bg-white/20 animate-pulse"></div>
            </div>
          </div>
          <div className="flex justify-between mt-2 text-sm text-gray-500">
            <span>{percentage}%</span>
            <span>Step {Math.ceil(importProgress.step)} of {importProgress.total}</span>
          </div>
        </div>
        
        {/* Step indicators */}
        <div className="flex gap-3 mt-4">
          {steps.map((step, i) => (
            <div 
              key={i}
              className={`flex flex-col items-center transition-all duration-300 ${
                i <= currentStepIndex ? 'opacity-100' : 'opacity-30'
              }`}
            >
              <div className={`w-10 h-10 rounded-full flex items-center justify-center text-lg
                ${i < currentStepIndex 
                  ? 'bg-green-500/20 border-2 border-green-500' 
                  : i === currentStepIndex 
                    ? 'bg-blue-500/20 border-2 border-blue-500 animate-pulse' 
                    : 'bg-gray-700/50 border-2 border-gray-600'
                }`}
              >
                {i < currentStepIndex ? '✓' : step.icon}
              </div>
              <span className={`text-xs mt-1 ${i <= currentStepIndex ? 'text-gray-300' : 'text-gray-600'}`}>
                {step.label}
              </span>
            </div>
          ))}
        </div>
        
        {/* Subtle hint */}
        <p className="text-gray-600 text-sm mt-8">
          This may take a moment for large presentations...
        </p>
      </div>
    );
  }

  if (state === 'presenting' && connection) {
    const autoLoginUrl = `${connection.url}?pin=${connection.pin}`;
    
    return (
      <div className="flex flex-col items-center justify-center h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900 text-white gap-6 p-8">
        <div className="absolute top-4 left-4 flex items-center gap-2">
          <div className="w-3 h-3 bg-green-500 rounded-full animate-pulse"></div>
          <span className="text-green-400 text-sm font-medium">Live</span>
        </div>
        
        <h1 className="text-3xl font-bold bg-gradient-to-r from-blue-400 to-purple-400 bg-clip-text text-transparent">
          Presenting
        </h1>
        <p className="text-slate-400">{presentation?.fileName}</p>

        <div className="bg-slate-800/50 backdrop-blur-sm rounded-2xl p-8 text-center border border-slate-700/50 shadow-2xl">
          {/* QR Code */}
          <div className="bg-white p-4 rounded-xl mb-6 inline-block">
            <QRCodeSVG 
              value={autoLoginUrl} 
              size={160}
              level="M"
              marginSize={0}
            />
          </div>
          
          <p className="text-slate-400 text-sm mb-1">Scan to connect instantly</p>
          <p className="text-slate-500 text-xs mb-6">or visit manually:</p>
          
          <p className="text-lg font-mono bg-slate-700/50 px-4 py-2 rounded-lg mb-6 border border-slate-600/50">
            {connection.url}
          </p>

          <p className="text-slate-400 mb-2">Enter this PIN:</p>
          <p className="text-5xl font-bold tracking-[0.3em] bg-gradient-to-r from-blue-400 to-cyan-400 bg-clip-text text-transparent">
            {connection.pin}
          </p>
        </div>

        <button
          onClick={handleStop}
          className="mt-4 px-8 py-3 bg-red-600/90 text-white rounded-xl hover:bg-red-500 transition-all text-lg font-medium shadow-lg shadow-red-900/30 hover:shadow-red-900/50"
        >
          Stop Presentation
        </button>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900">
      {/* Header */}
      <header className="flex justify-between items-center p-5 bg-slate-800/50 backdrop-blur-sm border-b border-slate-700/50">
        <div className="flex items-center gap-3">
          <img src={appIcon} alt="SlideCue" className="w-10 h-10 rounded-xl shadow-lg" />
          <div>
            <h1 className="text-xl font-bold text-white">SlideCue</h1>
            <p className="text-xs text-slate-400">PowerPoint Remote Control</p>
          </div>
        </div>
        <div className="flex gap-3">
          <button
            onClick={handleImport}
            className="px-5 py-2.5 bg-blue-600 text-white rounded-xl hover:bg-blue-500 transition-all font-medium shadow-lg shadow-blue-900/30 hover:shadow-blue-900/50 flex items-center gap-2"
          >
            <span>📁</span>
            Import PPTX
          </button>
          {state === 'loaded' && (
            <button
              onClick={handlePresent}
              className="px-5 py-2.5 bg-gradient-to-r from-green-600 to-emerald-600 text-white rounded-xl hover:from-green-500 hover:to-emerald-500 transition-all font-medium shadow-lg shadow-green-900/30 hover:shadow-green-900/50 flex items-center gap-2"
            >
              <span>▶️</span>
              Present
            </button>
          )}
        </div>
      </header>

      {error && (
        <div className="mx-6 mt-4 p-4 bg-red-500/10 border border-red-500/30 rounded-xl text-red-400 flex items-center gap-3">
          <span className="text-xl">⚠️</span>
          {error}
        </div>
      )}

      <main className="flex-1 p-6 overflow-auto">
        {presentation ? (
          <div>
            <div className="flex items-center gap-3 mb-6">
              <div className="w-12 h-12 bg-slate-700/50 rounded-xl flex items-center justify-center text-2xl">
                📄
              </div>
              <div>
                <p className="text-white font-medium">{presentation.fileName}</p>
                <p className="text-slate-400 text-sm">
                  {presentation.thumbnails.length} slide{presentation.thumbnails.length !== 1 ? 's' : ''}
                  {presentation.hiddenSlides.length > 0 && (
                    <span className="text-slate-500"> • {presentation.hiddenSlides.length} hidden</span>
                  )}
                </p>
              </div>
            </div>
            <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 xl:grid-cols-5 gap-4">
              {presentation.thumbnails.map((thumb, i) => (
                <div
                  key={i}
                  className="bg-slate-800/50 border border-slate-700/50 rounded-xl overflow-hidden shadow-lg hover:shadow-xl transition-all hover:border-slate-600/50 group"
                >
                  <div className="relative w-full bg-black flex items-center justify-center" style={{ minHeight: '120px', height: '150px' }}>
                    <img
                      src={`file://${thumb}`}
                      alt={`Slide ${i + 1}`}
                      className="max-w-full max-h-full object-contain"
                      onError={(e) => {
                        (e.target as HTMLImageElement).style.display = 'none';
                      }}
                    />
                    <div className="absolute inset-0 flex items-center justify-center bg-gradient-to-br from-slate-700 to-slate-800 -z-10">
                      <span className="text-4xl font-bold text-white/30">{i + 1}</span>
                    </div>
                  </div>
                  <div className="flex items-center justify-between px-3 py-2 bg-slate-800/80">
                    <span className="text-sm text-slate-400 font-medium">Slide {i + 1}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        ) : (
          <div className="flex-1 flex items-center justify-center h-full">
            <div className="border-2 border-dashed border-slate-700/50 rounded-2xl p-16 text-center bg-slate-800/20 hover:bg-slate-800/30 hover:border-slate-600/50 transition-all cursor-pointer group"
                 onClick={handleImport}>
              <img src={appIcon} alt="SlideCue" className="w-24 h-24 rounded-2xl mx-auto mb-6 group-hover:scale-110 transition-transform shadow-xl" />
              <h2 className="text-xl font-semibold text-white mb-2">Import a Presentation</h2>
              <p className="text-slate-400 text-sm mb-6 max-w-xs mx-auto">
                Click here or use the button above to import a PowerPoint file
              </p>
              <div className="inline-flex items-center gap-2 px-4 py-2 bg-slate-700/50 rounded-lg text-slate-400 text-sm">
                <span>📁</span>
                <span>.pptx, .ppt files supported</span>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
