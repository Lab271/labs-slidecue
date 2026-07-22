// SPDX-License-Identifier: Apache-2.0
// SPDX-FileCopyrightText: 2025-2026 Schuberg Philis / Lab271
import { useState, useEffect } from 'react';
import { QRCodeSVG } from 'qrcode.react';
import './types/electron.d.ts';
import appIcon from './assets/icon.png';
import sharepointQR from '../../../resources/sbp.link.sp_space.png';

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
  const [presentation, setPresentation] = useState<PresentationInfo | null>(null);
  const [connection, setConnection] = useState<ConnectionInfo | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [importProgress, setImportProgress] = useState<ImportProgress | null>(null);

  useEffect(() => {
    window.electronAPI.checkPowerPoint().then((installed) => {
      setState(installed ? 'idle' : 'no-powerpoint');
    });

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

  const handleStartSlideshow = async () => {
    try {
      await window.electronAPI.startSlideshow();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to start slideshow');
    }
  };

  // Loading state
  if (state === 'checking') {
    return (
      <div className="flex items-center justify-center h-screen bg-[#0f172a]">
        <div className="text-center">
          <img src={appIcon} alt="SlideCue" className="w-16 h-16 rounded-xl mx-auto mb-6 shadow-lg" />
          <div className="w-6 h-6 border-2 border-[#3b82f6]/30 border-t-[#3b82f6] rounded-full animate-spin mx-auto mb-4"></div>
          <p className="text-[#94a3b8]">Checking PowerPoint...</p>
        </div>
      </div>
    );
  }

  // No PowerPoint installed
  if (state === 'no-powerpoint') {
    return (
      <div className="flex flex-col items-center justify-center h-screen bg-[#0f172a] gap-6 p-8">
        <div className="w-16 h-16 bg-red-500/20 rounded-xl flex items-center justify-center border border-red-500/30">
          <svg className="w-8 h-8 text-red-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
          </svg>
        </div>
        <h1 className="text-xl font-bold text-red-400">PowerPoint Not Found</h1>
        <p className="text-[#94a3b8] text-center max-w-md">
          Microsoft PowerPoint is required to use SlideCue. Please install PowerPoint and restart the app.
        </p>
        <div className="px-4 py-2 bg-[#1e293b] rounded-lg border border-[#475569]">
          <p className="text-[#64748b] text-sm">Looking for: Microsoft PowerPoint.app</p>
        </div>
      </div>
    );
  }

  // Importing state
  if (state === 'importing' && importProgress) {
    const percentage = Math.round((importProgress.step / importProgress.total) * 100);
    
    return (
      <div className="flex flex-col items-center justify-center h-screen bg-[#0f172a] gap-8 p-8">
        <img src={appIcon} alt="SlideCue" className="w-16 h-16 rounded-xl shadow-lg" />
        
        <div className="text-center">
          <h1 className="text-2xl font-bold text-[#f8fafc] mb-2">Preparing Your Slides</h1>
          <p className="text-[#3b82f6]">{importProgress.message}</p>
        </div>
        
        <div className="w-80 max-w-full">
          <div className="bg-[#1e293b] rounded-full h-2 overflow-hidden border border-[#475569]">
            <div 
              className="bg-[#3b82f6] h-full transition-all duration-300"
              style={{ width: `${percentage}%` }}
            />
          </div>
          <div className="flex justify-between mt-2 text-sm text-[#64748b]">
            <span>{percentage}%</span>
            <span>Step {Math.ceil(importProgress.step)} of {importProgress.total}</span>
          </div>
        </div>
        
        <p className="text-[#64748b] text-sm">This may take a moment for large presentations...</p>
      </div>
    );
  }

  // Presenting state
  if (state === 'presenting' && connection) {
    const autoLoginUrl = `${connection.url}?pin=${connection.pin}`;
    
    return (
      <div className="flex flex-col items-center justify-center h-screen bg-[#0f172a] text-white gap-6 p-8">
        <div className="absolute top-4 left-4 flex items-center gap-2">
          <div className="w-2.5 h-2.5 bg-[#22c55e] rounded-full animate-pulse"></div>
          <span className="text-[#22c55e] text-sm font-medium">Live</span>
        </div>
        
        <h1 className="text-2xl font-bold text-[#f8fafc]">Presenting</h1>
        <p className="text-[#94a3b8]">{presentation?.fileName}</p>

        <div className="bg-[#1e293b] rounded-xl p-8 text-center border border-[#475569]">
          <div className="bg-white p-4 rounded-lg mb-6 inline-block">
            <QRCodeSVG 
              value={autoLoginUrl} 
              size={160}
              level="M"
              marginSize={0}
            />
          </div>
          
          <p className="text-[#94a3b8] text-sm mb-1">Scan to connect instantly</p>
          <p className="text-[#64748b] text-xs mb-6">or visit manually:</p>
          
          <p className="font-mono bg-[#334155] px-4 py-2 rounded-lg mb-6 border border-[#475569] text-[#f8fafc]">
            {connection.url}
          </p>

          <p className="text-[#94a3b8] mb-2">Enter PIN:</p>
          <p className="text-4xl font-bold tracking-[0.3em] text-[#3b82f6]">
            {connection.pin}
          </p>
        </div>

        <div className="flex gap-4">
          <button
            onClick={handleStartSlideshow}
            className="px-8 py-3 bg-[#3b82f6] text-white rounded-lg hover:bg-[#2563eb] transition-all font-medium"
          >
            Enter Presentation Mode
          </button>
          <button
            onClick={handleStop}
            className="px-8 py-3 bg-[#334155] text-white rounded-lg hover:bg-[#475569] transition-all font-medium border border-[#475569]"
          >
            Stop
          </button>
        </div>
      </div>
    );
  }

  // Main view (idle or loaded)
  return (
    <div className="flex flex-col h-screen bg-[#0f172a]">
      {/* Header */}
      <header className="flex justify-between items-center px-5 py-4 bg-[#1e293b] border-b border-[#475569]">
        <div className="flex items-center gap-3">
          <img src={appIcon} alt="SlideCue" className="w-10 h-10 rounded-lg" />
          <div>
            <h1 className="text-lg font-bold text-[#f8fafc]">SlideCue</h1>
            <p className="text-xs text-[#64748b]">PowerPoint Remote Control</p>
          </div>
        </div>
        <div className="h-10 flex items-center">
          {state === 'loaded' && (
            <button
              onClick={handleImport}
              className="px-4 py-2 bg-[#334155] text-[#f8fafc] rounded-lg hover:bg-[#475569] transition-all text-sm font-medium border border-[#475569]"
            >
              Load Different File
            </button>
          )}
        </div>
      </header>

      {error && (
        <div className="mx-4 mt-4 p-3 bg-red-500/10 border border-red-500/30 rounded-lg text-red-400 text-sm">
          {error}
        </div>
      )}

      <main className="flex-1 p-6 overflow-auto">
        {presentation ? (
          <div>
            {/* Loaded state card */}
            <div className="bg-[#1e293b] border border-[#22c55e]/30 rounded-xl p-6 mb-6">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-4">
                  <div className="w-14 h-14 bg-[#22c55e]/20 rounded-xl flex items-center justify-center border border-[#22c55e]/30">
                    <svg className="w-7 h-7 text-[#22c55e]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                    </svg>
                  </div>
                  <div>
                    <p className="text-[#22c55e] text-xs font-semibold mb-1">READY TO PRESENT</p>
                    <p className="text-[#f8fafc] font-bold text-xl mb-1">{presentation.fileName}</p>
                    <p className="text-[#94a3b8] text-sm">
                      {presentation.thumbnails.length} slides
                      {presentation.hiddenSlides.length > 0 && (
                        <span className="text-[#64748b]"> · {presentation.hiddenSlides.length} hidden</span>
                      )}
                    </p>
                  </div>
                </div>
                <button
                  onClick={handlePresent}
                  className="px-8 py-4 bg-[#22c55e] text-white rounded-xl hover:bg-[#16a34a] transition-all font-semibold text-lg"
                >
                  Start Remote Control
                </button>
              </div>
            </div>
            
            {/* Slide preview grid */}
            <div className="mb-4 flex items-center justify-between">
              <h3 className="text-[#94a3b8] font-medium">Slide Preview</h3>
              <p className="text-[#64748b] text-sm">{presentation.thumbnails.length} slides</p>
            </div>
            <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 xl:grid-cols-5 gap-3">
              {presentation.thumbnails.map((thumb, i) => (
                <div
                  key={i}
                  className="bg-[#1e293b] border border-[#475569] rounded-lg overflow-hidden hover:border-[#64748b] transition-all"
                >
                  <div className="relative bg-black aspect-video flex items-center justify-center">
                    <img
                      src={`file://${thumb}`}
                      alt={`Slide ${i + 1}`}
                      className="max-w-full max-h-full object-contain"
                      onError={(e) => {
                        (e.target as HTMLImageElement).style.display = 'none';
                      }}
                    />
                    <div className="absolute inset-0 flex items-center justify-center bg-[#334155] -z-10">
                      <span className="text-3xl font-bold text-[#475569]">{i + 1}</span>
                    </div>
                  </div>
                  <div className="px-3 py-2 bg-[#334155]">
                    <span className="text-xs text-[#94a3b8]">Slide {i + 1}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        ) : (
          /* Empty state - workflow steps */
          <div className="flex-1 flex items-center justify-center h-full">
            <div className="max-w-4xl w-full">
              <div className="text-center mb-10">
                <img src={appIcon} alt="SlideCue" className="w-16 h-16 rounded-xl mx-auto mb-4" />
                <h2 className="text-2xl font-bold text-[#f8fafc] mb-2">Get Started with SlideCue</h2>
                <p className="text-[#94a3b8]">Follow these steps to present your slides</p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                {/* Step 1 */}
                <div className="bg-[#1e293b] border border-[#475569] rounded-xl p-6 text-center hover:border-[#64748b] transition-all">
                  <div className="w-12 h-12 bg-[#3b82f6]/20 rounded-lg flex items-center justify-center mx-auto mb-4 border border-[#3b82f6]/30">
                    <span className="text-2xl font-bold text-[#3b82f6]">1</span>
                  </div>
                  <h3 className="text-lg font-semibold text-[#f8fafc] mb-2">Get Your PPTX</h3>
                  <p className="text-[#94a3b8] text-sm mb-4">
                    Scan to access presentations on SharePoint
                  </p>
                  <div className="bg-white p-3 rounded-lg inline-block">
                    <img src={sharepointQR} alt="SharePoint QR" className="w-28 h-28" />
                  </div>
                  <p className="text-xs text-[#64748b] mt-3">sbp.link/sp_space</p>
                </div>

                {/* Step 2 */}
                <div className="bg-[#1e293b] border border-[#475569] rounded-xl p-6 text-center hover:border-[#64748b] transition-all">
                  <div className="w-12 h-12 bg-[#8b5cf6]/20 rounded-lg flex items-center justify-center mx-auto mb-4 border border-[#8b5cf6]/30">
                    <span className="text-2xl font-bold text-[#8b5cf6]">2</span>
                  </div>
                  <h3 className="text-lg font-semibold text-[#f8fafc] mb-2">Load Presentation</h3>
                  <p className="text-[#94a3b8] text-sm mb-6">
                    Import your PowerPoint file
                  </p>
                  <button
                    onClick={handleImport}
                    className="px-6 py-3 bg-[#8b5cf6] text-white rounded-lg hover:bg-[#7c3aed] transition-all font-medium w-full"
                  >
                    Import PPTX
                  </button>
                </div>

                {/* Step 3 */}
                <div className="bg-[#1e293b] border border-[#475569] rounded-xl p-6 text-center hover:border-[#64748b] transition-all">
                  <div className="w-12 h-12 bg-[#22c55e]/20 rounded-lg flex items-center justify-center mx-auto mb-4 border border-[#22c55e]/30">
                    <span className="text-2xl font-bold text-[#22c55e]">3</span>
                  </div>
                  <h3 className="text-lg font-semibold text-[#f8fafc] mb-2">Start Presenting</h3>
                  <p className="text-[#94a3b8] text-sm mb-6">
                    Generate QR code for remote control
                  </p>
                  <div className="px-6 py-3 bg-[#334155] text-[#64748b] rounded-lg font-medium cursor-not-allowed">
                    Start Remote Control
                  </div>
                  <p className="text-xs text-[#64748b] mt-2">Import a file first</p>
                </div>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
