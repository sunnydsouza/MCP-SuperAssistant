import { resolve } from 'node:path';
import { defineConfig, type PluginOption } from 'vite';
import libAssetsPlugin from '@laynezh/vite-plugin-lib-assets';
import makeManifestPlugin from './utils/plugins/make-manifest-plugin.js';
import { watchPublicPlugin, watchRebuildPlugin } from '@extension/hmr';
import { watchOption } from '@extension/vite-config';
import env, { IS_DEV, IS_PROD } from '@extension/env';
import { nodePolyfills } from 'vite-plugin-node-polyfills';

const rootDir = resolve(import.meta.dirname);
const srcDir = resolve(rootDir, 'src');

const outDir = resolve(rootDir, '..', 'dist');
export default defineConfig({
  define: {
    'process.env': env,
  },
  envPrefix: ['VITE_', 'CEB_'],
  resolve: {
    // RegExp aliases require Vite's array form. These aliases keep Ajv runtime
    // code generation out of the Chrome/Edge Manifest V3 service worker.
    alias: [
      { find: '@root', replacement: rootDir },
      { find: '@src', replacement: srcDir },
      { find: '@assets', replacement: resolve(srcDir, 'assets') },
      { find: /^ajv-formats(\/.*)?$/, replacement: resolve(srcDir, 'shims', 'ajv-formats.ts') },
      { find: 'ajv/dist/compile/codegen', replacement: resolve(srcDir, 'shims', 'ajv-codegen.ts') },
      { find: /^ajv(\/.*)?$/, replacement: resolve(srcDir, 'shims', 'ajv.ts') },
    ],
  },
  plugins: [
    libAssetsPlugin({
      outputPath: outDir,
    }) as PluginOption,
    watchPublicPlugin(),
    makeManifestPlugin({ outDir }),
    IS_DEV && watchRebuildPlugin({ reload: true, id: 'chrome-extension-hmr' }),
    nodePolyfills(),
  ],
  publicDir: resolve(rootDir, 'public'),
  build: {
    lib: {
      name: 'BackgroundScript',
      fileName: 'background',
      formats: ['es'],
      entry: resolve(srcDir, 'background', 'index.ts'),
    },
    outDir,
    emptyOutDir: false,
    sourcemap: IS_DEV,
    minify: IS_PROD,
    reportCompressedSize: IS_PROD,
    watch: watchOption,
    rollupOptions: {
      external: ['chrome'],
    },
  },
});
