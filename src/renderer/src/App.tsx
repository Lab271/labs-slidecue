import { useState, useEffect } from 'react';
import './types/electron.d.ts';

type AppState = 'checking' | 'no-powerpoint' | 'idle' | 'loaded' | 'presenting';

interface PresentationInfo {
  filePath: string;
  fileName: string;
  thumbnails: string[];
}

interface ConnectionInfo {
  url: string;
  pin: string;
}

export default function App() {
  const [state, setState] = useState<AppState>('checking');
  const [presentation, setPresentation] = useState<PresentationInfo | null>(
    null
  );
  const [connection, setConnection] = useState<ConnectionInfo | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    window.electronAPI.checkPowerPoint().then((installed) => {
      setState(installed ? 'idle' : 'no-powerpoint');
    });
  }, []);

  const handleImport = async () => {
    try {
      setError(null);
      const result = await window.electronAPI.importPresentation();
      if (result) {
        setPresentation(result);
        setState('loaded');
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to import');
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
              {presentation.fileName} • {presentation.thumbnails.length} slides
            </p>
            <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 xl:grid-cols-5 gap-4">
              {presentation.thumbnails.map((thumb, i) => (
                <div
                  key={i}
                  className="bg-white border rounded-lg overflow-hidden shadow-sm hover:shadow-md transition-shadow"
                >
                  <img
                    src={`slidecue://${thumb}`}
                    alt={`Slide ${i + 1}`}
                    className="w-full aspect-video object-contain bg-gray-100"
                    onError={(e) => {
                      (e.target as HTMLImageElement).style.display = 'none';
                      (e.target as HTMLImageElement).nextElementSibling?.classList.remove('hidden');
                    }}
                  />
                  <div className="hidden w-full aspect-video flex items-center justify-center bg-gradient-to-br from-gray-700 to-gray-800">
                    <span className="text-4xl font-bold text-white/70">{i + 1}</span>
                  </div>
                  <p className="text-center py-2 text-sm text-gray-600 bg-gray-50">
                    Slide {i + 1}
                  </p>
                </div>
              ))}
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
