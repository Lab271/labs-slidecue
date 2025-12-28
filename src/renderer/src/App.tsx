import { useState, useEffect } from 'react';
import './types/electron.d.ts';

type AppState = 'checking' | 'no-powerpoint' | 'idle' | 'importing' | 'loaded' | 'presenting';

interface PresentationInfo {
  filePath: string;
  fileName: string;
  thumbnails: string[];
  totalSlides: number;
  visibleSlideCount: number;
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
      <div className="flex items-center justify-center h-screen bg-gray-50">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-500 mx-auto mb-4"></div>
          <p className="text-xl text-gray-600">Checking PowerPoint...</p>
        </div>
      </div>
    );
  }

  if (state === 'no-powerpoint') {
    return (
      <div className="flex flex-col items-center justify-center h-screen bg-gray-50 gap-4 p-8">
        <div className="text-6xl">⚠️</div>
        <h1 className="text-2xl font-bold text-red-600">PowerPoint Not Found</h1>
        <p className="text-gray-600 text-center max-w-md">
          Microsoft PowerPoint is required to use SlideCue. Please install
          PowerPoint and restart the app.
        </p>
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
    return (
      <div className="flex flex-col items-center justify-center h-screen bg-gray-900 text-white gap-6 p-8">
        <h1 className="text-3xl font-bold">Presenting</h1>
        <p className="text-gray-400">{presentation?.fileName}</p>

        <div className="bg-gray-800 rounded-xl p-8 text-center">
          <p className="text-gray-400 mb-2">Connect from any device:</p>
          <p className="text-xl font-mono bg-gray-700 px-4 py-2 rounded mb-6">
            {connection.url}
          </p>

          <p className="text-gray-400 mb-2">Enter this PIN:</p>
          <p className="text-6xl font-bold tracking-[0.3em] text-blue-400">
            {connection.pin}
          </p>
        </div>

        <button
          onClick={handleStop}
          className="mt-8 px-8 py-3 bg-red-600 text-white rounded-lg hover:bg-red-700 transition-colors text-lg font-medium"
        >
          Stop Presentation
        </button>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-screen bg-gray-50">
      <header className="flex justify-between items-center p-6 bg-white border-b">
        <h1 className="text-2xl font-bold text-gray-800">SlideCue</h1>
        <div className="flex gap-3">
          <button
            onClick={handleImport}
            className="px-4 py-2 bg-blue-500 text-white rounded-lg hover:bg-blue-600 transition-colors font-medium"
          >
            Import PPTX
          </button>
          {state === 'loaded' && (
            <button
              onClick={handlePresent}
              className="px-4 py-2 bg-green-500 text-white rounded-lg hover:bg-green-600 transition-colors font-medium"
            >
              Present
            </button>
          )}
        </div>
      </header>

      {error && (
        <div className="mx-6 mt-4 p-4 bg-red-100 border border-red-300 rounded-lg text-red-700">
          {error}
        </div>
      )}

      <main className="flex-1 p-6 overflow-auto">
        {presentation ? (
          <div>
            <p className="text-gray-600 mb-4">
              {presentation.fileName} • {presentation.visibleSlideCount} visible slide{presentation.visibleSlideCount !== 1 ? 's' : ''}
              {presentation.hiddenSlides.length > 0 && (
                <span className="text-gray-400"> ({presentation.hiddenSlides.length} hidden)</span>
              )}
            </p>
            <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 xl:grid-cols-5 gap-4">
              {presentation.thumbnails.map((thumb, i) => {
                const isHidden = presentation.hiddenSlides.includes(i + 1);
                return (
                  <div
                    key={i}
                    className={`bg-white border rounded-lg overflow-hidden shadow-sm hover:shadow-md transition-shadow ${isHidden ? 'opacity-50' : ''}`}
                  >
                    <div className="relative w-full aspect-video bg-gray-100">
                      <img
                        src={`file://${thumb}`}
                        alt={`Slide ${i + 1}`}
                        className="w-full h-full object-contain"
                        onError={(e) => {
                          (e.target as HTMLImageElement).style.display = 'none';
                        }}
                      />
                      <div className="absolute inset-0 flex items-center justify-center bg-gradient-to-br from-gray-700 to-gray-800 -z-10">
                        <span className="text-4xl font-bold text-white/70">{i + 1}</span>
                      </div>
                      {isHidden && (
                        <div className="absolute top-2 right-2 bg-gray-900/80 text-white text-xs px-2 py-1 rounded flex items-center gap-1">
                          <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13.875 18.825A10.05 10.05 0 0112 19c-4.478 0-8.268-2.943-9.543-7a9.97 9.97 0 011.563-3.029m5.858.908a3 3 0 114.243 4.243M9.878 9.878l4.242 4.242M9.88 9.88l-3.29-3.29m7.532 7.532l3.29 3.29M3 3l3.59 3.59m0 0A9.953 9.953 0 0112 5c4.478 0 8.268 2.943 9.543 7a10.025 10.025 0 01-4.132 5.411m0 0L21 21" />
                          </svg>
                          Hidden
                        </div>
                      )}
                    </div>
                    <p className="text-center py-2 text-sm text-gray-600 bg-gray-50">
                      Slide {i + 1}{isHidden && ' (hidden)'}
                    </p>
                  </div>
                );
              })}
            </div>
          </div>
        ) : (
          <div className="flex-1 flex items-center justify-center h-full">
            <div className="border-2 border-dashed border-gray-300 rounded-xl p-16 text-center">
              <div className="text-5xl mb-4">📄</div>
              <p className="text-gray-500 text-lg">
                Import a PowerPoint file to get started
              </p>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
