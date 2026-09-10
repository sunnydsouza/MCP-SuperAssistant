import { pluginRegistry } from './plugin-registry';
import { Microsoft365CopilotAdapter } from './adapters/m365copilot.adapter';
import { createLogger } from '@extension/shared/lib/logger';

const logger = createLogger('Microsoft365CopilotRegistration');

/**
 * Register Microsoft Copilot after the core registry has initialized.
 *
 * Microsoft currently surfaces Copilot Chat on both copilot.cloud.microsoft
 * and m365.cloud.microsoft depending on entry point and tenant routing.
 */
export async function initializeMicrosoft365CopilotSupport(): Promise<void> {
  const hostname = window.location.hostname.toLowerCase();
  const isMicrosoftCopilot =
    hostname === 'copilot.cloud.microsoft' ||
    hostname === 'm365.cloud.microsoft' ||
    hostname.endsWith('.copilot.cloud.microsoft') ||
    hostname.endsWith('.m365.cloud.microsoft');

  if (!isMicrosoftCopilot) return;

  const pluginName = 'Microsoft365CopilotAdapter';

  if (!pluginRegistry.isPluginRegistered(pluginName)) {
    await pluginRegistry.register(new Microsoft365CopilotAdapter(), {
      id: 'm365-copilot-adapter',
      name: 'Microsoft Copilot Adapter',
      description: 'Adapter for Microsoft Copilot Chat on copilot.cloud.microsoft and m365.cloud.microsoft',
      version: '1.0.1',
      enabled: true,
      priority: 10,
      settings: {
        logLevel: 'info',
      },
    });
    logger.debug('Microsoft Copilot adapter registered');
  }

  await pluginRegistry.activatePlugin(pluginName);
  logger.debug(`Microsoft Copilot adapter activated for ${hostname}`);
}
