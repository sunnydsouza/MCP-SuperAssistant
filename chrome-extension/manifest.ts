import { readFileSync } from 'node:fs';

const packageJson = JSON.parse(readFileSync('./package.json', 'utf8'));

/**
 * @prop default_locale
 * if you want to support multiple languages, you can use the following reference
 * https://developer.mozilla.org/en-US/docs/Mozilla/Add-ons/WebExtensions/Internationalization
 *
 * @prop browser_specific_settings
 * Must be unique to your extension to upload to addons.mozilla.org
 * (you can delete if you only want a chrome extension)
 *
 * @prop permissions
 * Firefox doesn't support sidePanel (It will be deleted in manifest parser)
 *
 * @prop content_scripts
 * css: ['content.css'], // public folder
 */
const manifest = {
  manifest_version: 3,
  default_locale: 'en',
  name: 'MCP SuperAssistant',
  browser_specific_settings: {
    gecko: {
      id: 'saurabh@mcpsuperassistant.ai',
    },
  },
  version: packageJson.version,
  description: 'MCP SuperAssistant',
  host_permissions: [
    '*://*.perplexity.ai/*',
    '*://*.chat.openai.com/*',
    '*://*.chatgpt.com/*',
    '*://*.grok.com/*',
    '*://*.x.com/*',
    '*://*.twitter.com/*',
    '*://*.gemini.google.com/*',
    '*://*.aistudio.google.com/*',
    '*://*.openrouter.ai/*',
    '*://*.google-analytics.com/*',
    '*://*.chat.deepseek.com/*',
    '*://*.t3.chat/*',
    '*://*.chat.mistral.ai/*',
    '*://*.github.com/*',
    '*://*.copilot.github.com/*',
    '*://*.kimi.com/*',
    '*://*.chat.z.ai/*',
    '*://*.chat.qwen.ai/*',
    '*://m365.cloud.microsoft/*',
    '*://copilot.cloud.microsoft/*',
  ],

  permissions: ['storage', 'clipboardWrite'],
  background: {
    service_worker: 'background.js',
    type: 'module',
  },
  icons: {
    128: 'icon-128.png',
    34: 'icon-34.png',
  },
  content_scripts: [
    {
      matches: ['*://*.perplexity.ai/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    {
      matches: ['*://*.chat.openai.com/*', '*://*.chatgpt.com/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    {
      matches: ['*://*.grok.com/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    {
      matches: ['*://*.x.com/*', '*://*.twitter.com/*', '*://*.x.com/i/grok*', '*://*.twitter.com/i/grok*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    {
      matches: ['*://*.gemini.google.com/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    {
      matches: ['*://*.aistudio.google.com/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    {
      matches: ['*://*.openrouter.ai/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    {
      matches: ['*://*.chat.deepseek.com/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    {
      matches: ['*://*.kagi.com/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    {
      matches: ['*://*.t3.chat/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    {
      matches: ['*://*.chat.mistral.ai/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    {
      matches: ['*://*.github.com/*', '*://*.copilot.github.com/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    {
      matches: ['*://*.kimi.com/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    {
      matches: ['*://*.chat.z.ai/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    {
      matches: ['*://*.chat.qwen.ai/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
    // Microsoft Copilot app / Copilot Chat can currently surface on either
    // copilot.cloud.microsoft or m365.cloud.microsoft depending on entry point.
    {
      matches: ['*://m365.cloud.microsoft/*', '*://copilot.cloud.microsoft/*'],
      js: ['content/index.iife.js'],
      run_at: 'document_idle',
    },
  ],
  web_accessible_resources: [
    {
      resources: ['*.js', '*.css', 'content/*.css', '*.svg', 'icon-128.png', 'icon-34.png'],
      matches: ['*://*/*'],
    },
  ],
} satisfies chrome.runtime.ManifestV3;

export default manifest;
