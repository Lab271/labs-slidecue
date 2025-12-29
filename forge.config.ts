import type { ForgeConfig } from '@electron-forge/shared-types';
import { MakerSquirrel } from '@electron-forge/maker-squirrel';
import { MakerDMG } from '@electron-forge/maker-dmg';
import { MakerZIP } from '@electron-forge/maker-zip';

const config: ForgeConfig = {
  packagerConfig: {
    icon: './resources/icon',
    appBundleId: 'com.lab271.slidecue',
    asar: true,
    extraResource: ['./out/remote'],
    // Don't use .gitignore patterns - explicitly define what to ignore
    ignore: (file) => {
      if (!file) return false;
      // Include these paths
      if (
        file.startsWith('/out') ||
        file.startsWith('/package.json') ||
        file.startsWith('/node_modules')
      ) {
        return false;
      }
      // Ignore everything else (src, resources, etc.)
      if (
        file.startsWith('/src') ||
        file.startsWith('/resources') ||
        file.startsWith('/.') ||
        file.startsWith('/forge.config') ||
        file.startsWith('/electron.vite') ||
        file.startsWith('/tsconfig') ||
        file.startsWith('/README') ||
        file.startsWith('/LICENSE') ||
        file.startsWith('/TODO') ||
        file.startsWith('/ARCHITECTURE') ||
        file.startsWith('/IMPLEMENTATION')
      ) {
        return true;
      }
      return false;
    },
  },
  makers: [
    new MakerSquirrel({
      name: 'SlideCue',
      setupIcon: './resources/icon.ico',
    }),
    new MakerDMG({
      name: 'SlideCue',
      icon: './resources/icon.icns',
    }),
    new MakerZIP({}, ['darwin']),
  ],
  publishers: [
    {
      name: '@electron-forge/publisher-github',
      config: {
        repository: {
          owner: 'LAB271',
          name: 'SlideCue',
        },
        prerelease: false,
      },
    },
  ],
};

export default config;
