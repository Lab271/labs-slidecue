interface ElectronAPI {
  checkPowerPoint: () => Promise<boolean>;
  importPresentation: () => Promise<{
    filePath: string;
    fileName: string;
    thumbnails: string[];
    totalSlides: number;
  } | null>;
  startPresentation: () => Promise<{ url: string; pin: string }>;
  stopPresentation: () => Promise<void>;
  getSlideInfo: () => Promise<{ current: number; total: number }>;
  onImportProgress: (callback: (data: { step: number; total: number; message: string }) => void) => () => void;
}

declare global {
  interface Window {
    electronAPI: ElectronAPI;
  }
}

export {};
