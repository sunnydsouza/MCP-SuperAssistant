import { pluginRegistry } from './plugin-registry';
import { Microsoft365CopilotAdapter } from './adapters/m365copilot.adapter';
import { createLogger } from '@extension/shared/lib/logger';

const logger = createLogger('Microsoft365CopilotRegistration');

/**
 * Register Microsoft 365 Copilot after the core registry has initialized.
 *
 * The application initializes built-in adapter factories before services. Keeping
 * this registration here makes the M365 addition isolated from the large built-in
 * registry while still using the same lazy factory/activation mechanism.
 */
export async function initializeMicrosoft365CopilotSupport(): Promise<void> {
  const factoryName = 'm365-copilot-adapter';

  if (!pluginRegistry.isFactoryRegistered(factoryName) && !pluginRegistry.isPluginRegistered('Microsoft365CopilotAdapter')) {
    pluginRegistry.registerAdapterFactory({
      name: factoryName,
      version: '1.0.0',
      type: 'website-adapter',
      hostnames: ['m365.cloud.microsoft'],
      capabilities: ['text-insertion', 'form-submission', 'dom-manipulation'],
      create: () => new Microsoft365CopilotAdapter(),
      config: {
        id: factoryName,
        name: 'Microsoft 365 Copilot Adapter',
        description: 'Adapter for Microsoft 365 Copilot Chat on m365.cloud.microsoft',
        version: '1.0.0',
        enabled: true,
        priority: 10,
        settings: {
          logLevel: 'info',
        },
      },
    });

    logger.debug('Microsoft 365 Copilot adapter factory registered');
  }

  if (window.location.hostname === 'm365.cloud.microsoft' || window.location.hostname.endsWith('.m365.cloud.microsoft')) {
    await pluginRegistry.activatePluginForHostname(window.location.hostname);
    logger.debug('Microsoft 365 Copilot adapter activation requested for current page');
  }
}
