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
    // Disable default .gitignore-based ignoring
    // Include everything, package will still work
    ignore: () => false,
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
          owner: 'Lab271',
          name: 'labs-slidecue',
        },
        prerelease: false,
      },
    },
  ],
};

export default config;
