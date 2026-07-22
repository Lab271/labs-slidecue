// SPDX-License-Identifier: Apache-2.0
// SPDX-FileCopyrightText: 2025-2026 Schuberg Philis / Lab271
import { useState, useEffect, useRef } from 'react';
import { io, Socket } from 'socket.io-client';

type ViewMode = 'current' | 'next';

interface SlideInfo {
  currentSlide: number;
  totalSlides: number;
  nextVisibleSlide: number | null;
  animationStep: number;
  animationsOnSlide: number;
  isLastSlide: boolean;
  currentNotes: string;
  nextNotes: string;
}

export default function App() {
  const [pin, setPin] = useState('');
  const [authenticated, setAuthenticated] = useState(false);
  const [error, setError] = useState('');
  const [connecting, setConnecting] = useState(false);
  const [slideInfo, setSlideInfo] = useState<SlideInfo>({
    currentSlide: 1,
    totalSlides: 1,
    nextVisibleSlide: null,
    animationStep: 0,
    animationsOnSlide: 0,
    isLastSlide: false,
    currentNotes: '',
    nextNotes: ''
  });
  const [viewMode, setViewMode] = useState<ViewMode>('next');
  const socketRef = useRef<Socket | null>(null);
  const inputRefs = useRef<(HTMLInputElement | null)[]>([]);

  useEffect(() => {
    return () => {
      socketRef.current?.close();
    };
  }, []);

  const handlePinChange = (index: number, value: string) => {
    if (!/^\d*$/.test(value)) return;
    
    const newPin = pin.split('');
    newPin[index] = value;
    const updatedPin = newPin.join('').slice(0, 4);
    setPin(updatedPin);
    
    // Auto-focus next input
    if (value && index < 3) {
      inputRefs.current[index + 1]?.focus();
    }
  };

  const handleKeyDown = (index: number, e: React.KeyboardEvent) => {
    if (e.key === 'Backspace' && !pin[index] && index > 0) {
      inputRefs.current[index - 1]?.focus();
    }
  };

  const handleConnect = () => {
    if (pin.length !== 4) return;
    
    setConnecting(true);
    setError('');

    const socket = io(window.location.origin, {
      auth: { pin },
    });

    socket.on('connect', () => {
      setAuthenticated(true);
      setConnecting(false);
      setError('');
    });

    socket.on('connect_error', () => {
      setError('Invalid PIN');
      setConnecting(false);
      socket.close();
    });

    socket.on('slide-changed', (info: SlideInfo) => {
      setSlideInfo(info);
    });

    socket.on('disconnect', () => {
      setAuthenticated(false);
      setError('Disconnected from presentation');
    });

    socketRef.current = socket;
  };

  const handleNext = () => socketRef.current?.emit('next');
  const handlePrev = () => socketRef.current?.emit('prev');

  if (!authenticated) {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen p-6 bg-gray-100">
        <div className="bg-white rounded-2xl shadow-lg p-8 w-full max-w-sm">
          <h1 className="text-2xl font-bold text-center mb-2">SlideCue</h1>
          <p className="text-gray-500 text-center mb-8">Enter the 4-digit PIN</p>
          
          <div className="flex justify-center gap-3 mb-6">
            {[0, 1, 2, 3].map((i) => (
              <input
                key={i}
                ref={(el) => { inputRefs.current[i] = el; }}
                type="text"
                inputMode="numeric"
                maxLength={1}
                value={pin[i] || ''}
                onChange={(e) => handlePinChange(i, e.target.value)}
                onKeyDown={(e) => handleKeyDown(i, e)}
                className="w-14 h-16 text-center text-3xl font-bold border-2 rounded-xl focus:border-blue-500 focus:outline-none transition-colors"
                autoFocus={i === 0}
              />
            ))}
          </div>

          {error && (
            <p className="text-red-500 text-center mb-4">{error}</p>
          )}

          <button
            onClick={handleConnect}
            disabled={pin.length !== 4 || connecting}
            className="w-full py-4 bg-blue-500 text-white rounded-xl text-lg font-semibold disabled:opacity-50 disabled:cursor-not-allowed hover:bg-blue-600 transition-colors"
          >
            {connecting ? 'Connecting...' : 'Connect'}
          </button>
        </div>
      </div>
    );
  }

  const previewSlide =
    viewMode === 'current'
      ? slideInfo.currentSlide
      : (slideInfo.nextVisibleSlide || slideInfo.currentSlide);

  return (
    <div className="flex flex-col h-screen bg-gray-100">
      {/* Slide Preview */}
      <div className="flex-1 p-4 overflow-hidden">
        <div className="bg-white rounded-xl shadow-lg overflow-hidden h-full flex flex-col">
          <img
            src={`/thumbnails/slide_${previewSlide.toString().padStart(3, '0')}.png`}
            alt={`Slide ${previewSlide}`}
            className="w-full flex-1 object-contain bg-black"
          />
        </div>
      </div>

      {/* View Toggle */}
      <div className="flex justify-center gap-2 px-4">
        <button
          onClick={() => setViewMode('current')}
          className={`px-4 py-2 rounded-lg font-medium transition-colors ${
            viewMode === 'current'
              ? 'bg-blue-500 text-white'
              : 'bg-white text-gray-600'
          }`}
        >
          Current
        </button>
        <button
          onClick={() => setViewMode('next')}
          className={`px-4 py-2 rounded-lg font-medium transition-colors ${
            viewMode === 'next'
              ? 'bg-blue-500 text-white'
              : 'bg-white text-gray-600'
          }`}
        >
          Next
        </button>
      </div>

      {/* Slide Counter */}
      <div className="text-center py-3">
        <span className="text-2xl font-bold text-gray-800">
          {slideInfo.currentSlide} / {slideInfo.totalSlides}
        </span>
      </div>

      {/* Navigation Buttons */}
      <div className="grid grid-cols-2 gap-3 p-4 pb-8">
        <button
          onClick={handlePrev}
          disabled={slideInfo.currentSlide <= 1}
          className="py-6 bg-gray-800 text-white text-xl font-semibold rounded-xl disabled:opacity-30 active:bg-gray-700 transition-colors"
        >
          ← Previous
        </button>
        <button
          onClick={handleNext}
          disabled={slideInfo.isLastSlide}
          className="py-6 bg-blue-600 text-white text-xl font-semibold rounded-xl disabled:opacity-30 active:bg-blue-700 transition-colors"
        >
          Next →
        </button>
      </div>
    </div>
  );
}
