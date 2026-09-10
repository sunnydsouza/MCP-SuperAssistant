import { pluginRegistry } from './plugin-registry';
import { Microsoft365CopilotAdapter } from './adapters/m365copilot.adapter';
import { addMicrosoftCopilotBridgeInstructions } from './m365copilot-instructions';
import { createLogger } from '@extension/shared/lib/logger';

const logger = createLogger('Microsoft365CopilotRegistration');

/**
 * Microsoft Copilot distinguishes its native tool registry from capabilities
 * described in ordinary chat text. MCP SuperAssistant is an external browser
 * bridge, so make that distinction explicit only when inserting the generated
 * SuperAssistant instruction prompt. Function results and normal text are left
 * untouched.
 */
class Microsoft365CopilotBridgeAdapter extends Microsoft365CopilotAdapter {
  override async insertText(text: string, options?: { targetElement?: HTMLElement }): Promise<boolean> {
    return super.insertText(addMicrosoftCopilotBridgeInstructions(text), options);
  }
}

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
    await pluginRegistry.register(new Microsoft365CopilotBridgeAdapter(), {
      id: 'm365-copilot-adapter',
      name: 'Microsoft Copilot Adapter',
      description: 'Adapter for Microsoft Copilot Chat on copilot.cloud.microsoft and m365.cloud.microsoft',
      version: '1.1.1',
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
