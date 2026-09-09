import { pluginRegistry } from './plugin-registry';
import { Microsoft365CopilotAdapter } from './adapters/m365copilot.adapter';
import { createLogger } from '@extension/shared/lib/logger';

const logger = createLogger('Microsoft365CopilotRegistration');

/**
 * Register Microsoft 365 Copilot after the core registry has initialized.
 *
 * The content script is enabled for M365 by the extension manifest, and services
 * initialize after the core plugin registry. Registering the concrete adapter here
 * keeps the M365 change isolated from the large built-in registry and avoids doing
 * any M365-specific work on other supported sites.
 */
export async function initializeMicrosoft365CopilotSupport(): Promise<void> {
  const hostname = window.location.hostname.toLowerCase();
  const isM365Copilot = hostname === 'm365.cloud.microsoft' || hostname.endsWith('.m365.cloud.microsoft');
  if (!isM365Copilot) return;

  const pluginName = 'Microsoft365CopilotAdapter';

  if (!pluginRegistry.isPluginRegistered(pluginName)) {
    await pluginRegistry.register(new Microsoft365CopilotAdapter(), {
      id: 'm365-copilot-adapter',
      name: 'Microsoft 365 Copilot Adapter',
      description: 'Adapter for Microsoft 365 Copilot Chat on m365.cloud.microsoft',
      version: '1.0.0',
      enabled: true,
      priority: 10,
      settings: {
        logLevel: 'info',
      },
    });
    logger.debug('Microsoft 365 Copilot adapter registered');
  }

  await pluginRegistry.activatePlugin(pluginName);
  logger.debug('Microsoft 365 Copilot adapter activated for current page');
}
