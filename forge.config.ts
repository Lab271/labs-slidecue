import type { ForgeConfig } from '@electron-forge/shared-types';
import { MakerSquirrel } from '@electron-forge/maker-squirrel';

const config: ForgeConfig = {
  packagerConfig: {
    icon: './resources/icon',
    appBundleId: 'com.slidecue.app',
    asar: true,
    extraResource: ['./out/remote'],
  },
  makers: [
    new MakerSquirrel({
      name: 'SlideCue',
      setupIcon: './resources/icon.ico',
    }),
  ],
  publishers: [
    {
      name: '@electron-forge/publisher-github',
      config: {
        repository: {
          owner: 'yourusername',
          name: 'slidecue',
        },
        prerelease: false,
      },
    },
  ],
};

export default config;
