import { pluginRegistry } from './plugin-registry';
import { Microsoft365CopilotAdapter } from './adapters/m365copilot.adapter';
import { addMicrosoftCopilotBridgeInstructions } from './m365copilot-instructions';
import { initializeMicrosoftCopilotFunctionBridge } from './m365copilot-function-bridge';
import { createLogger } from '@extension/shared/lib/logger';

const logger = createLogger('Microsoft365CopilotRegistration');

const FILE_INPUT_SELECTOR = 'input[type="file"]';
const PROMPT_SELECTOR = [
  'textarea[placeholder*="message" i]',
  'textarea[aria-label*="message" i]',
  '[contenteditable="true"][role="textbox"]',
  '[contenteditable="true"][data-lexical-editor="true"]',
  '[contenteditable="true"][aria-label*="message" i]',
  '[contenteditable="true"][aria-label*="prompt" i]',
].join(',');

const sleep = (milliseconds: number): Promise<void> =>
  new Promise(resolve => setTimeout(resolve, milliseconds));

/**
 * Prefer Microsoft's real file input when it is present. Hidden file inputs are
 * intentionally included: modern web apps commonly keep the input hidden and
 * trigger it from the composer '+' menu.
 */
function findMicrosoftFileInput(): HTMLInputElement | null {
  const inputs = Array.from(document.querySelectorAll<HTMLInputElement>(FILE_INPUT_SELECTOR)).filter(
    input => !input.disabled && input.getAttribute('aria-disabled') !== 'true',
  );

  if (inputs.length === 0) return null;

  // Prefer an input whose accept list can handle plain text, then the last input
  // in DOM order because it is usually associated with the active composer.
  const textInput = inputs.find(input => {
    const accept = (input.accept || '').toLowerCase();
    return !accept || accept.includes('text') || accept.includes('.txt') || accept.includes('*/*');
  });

  return textInput ?? inputs[inputs.length - 1] ?? null;
}

/**
 * Microsoft may not create its hidden file input until the Add/Attach control is
 * opened. Open that control and briefly wait for the input to be mounted.
 */
async function revealMicrosoftFileInput(): Promise<HTMLInputElement | null> {
  const existing = findMicrosoftFileInput();
  if (existing) return existing;

  const prompt = document.querySelector<HTMLElement>(PROMPT_SELECTOR);
  const promptRect = prompt?.getBoundingClientRect();

  const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>('button')).filter(button => {
    if (button.disabled || button.getAttribute('aria-disabled') === 'true') return false;
    const style = window.getComputedStyle(button);
    const rect = button.getBoundingClientRect();
    if (style.display === 'none' || style.visibility === 'hidden' || rect.width === 0 || rect.height === 0) return false;

    const label = `${button.getAttribute('aria-label') ?? ''} ${button.getAttribute('title') ?? ''} ${button.textContent ?? ''}`.trim();
    const looksLikeAttachmentControl = /\b(add|attach|upload|plus)\b/i.test(label) || button.textContent?.trim() === '+';
    if (!looksLikeAttachmentControl) return false;

    if (!promptRect) return true;
    const centerY = rect.top + rect.height / 2;
    const promptCenterY = promptRect.top + promptRect.height / 2;
    return Math.abs(centerY - promptCenterY) <= 140;
  });

  const attachButton = buttons[0];
  if (!attachButton) return null;

  attachButton.click();

  for (let attempt = 0; attempt < 12; attempt += 1) {
    await sleep(75);
    const input = findMicrosoftFileInput();
    if (input) return input;
  }

  return null;
}

function assignFileToInput(input: HTMLInputElement, file: File): boolean {
  try {
    const transfer = new DataTransfer();
    transfer.items.add(file);
    input.files = transfer.files;

    input.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
    input.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
    return true;
  } catch (error) {
    logger.warn('Failed to assign instructions to Microsoft Copilot file input:', error);
    return false;
  }
}

/**
 * Fallback for Microsoft builds that accept composer drops but do not expose a
 * usable file input to the page. The browser extension supplies the File through
 * DataTransfer just as a user drag/drop operation would.
 */
function dropFileOnMicrosoftComposer(file: File): boolean {
  const prompt = document.querySelector<HTMLElement>(PROMPT_SELECTOR);
  if (!prompt) return false;

  const target =
    (prompt.closest('form') as HTMLElement | null) ??
    prompt.parentElement ??
    prompt;

  try {
    const transfer = new DataTransfer();
    transfer.items.add(file);

    for (const type of ['dragenter', 'dragover', 'drop'] as const) {
      target.dispatchEvent(
        new DragEvent(type, {
          bubbles: true,
          cancelable: true,
          composed: true,
          dataTransfer: transfer,
        }),
      );
    }
    return true;
  } catch (error) {
    logger.warn('Failed to simulate Microsoft Copilot file drop:', error);
    return false;
  }
}

/**
 * Microsoft Copilot distinguishes its native tool registry from capabilities
 * described in ordinary chat text. MCP SuperAssistant is an external browser
 * bridge, so make that distinction explicit both for inline instructions and
 * for instructions attached as a file.
 */
class Microsoft365CopilotBridgeAdapter extends Microsoft365CopilotAdapter {
  override async insertText(text: string, options?: { targetElement?: HTMLElement }): Promise<boolean> {
    return super.insertText(addMicrosoftCopilotBridgeInstructions(text), options);
  }

  override async attachFile(file: File, options?: { inputElement?: HTMLInputElement }): Promise<boolean> {
    try {
      const originalText = await file.text();
      const bridgedText = addMicrosoftCopilotBridgeInstructions(originalText);
      const instructionsFile = new File([bridgedText], 'mcp_superassistant_instructions.txt', {
        type: 'text/plain',
        lastModified: Date.now(),
      });

      let input = options?.inputElement ?? findMicrosoftFileInput();
      if (!input) input = await revealMicrosoftFileInput();

      if (input && assignFileToInput(input, instructionsFile)) {
        logger.debug('Attached MCP instructions to Microsoft Copilot via file input');
        return true;
      }

      const dropped = dropFileOnMicrosoftComposer(instructionsFile);
      if (dropped) {
        logger.debug('Attached MCP instructions to Microsoft Copilot via drag/drop fallback');
        return true;
      }

      logger.warn('Microsoft Copilot file attachment failed: no usable file input or drop target');
      return false;
    } catch (error) {
      logger.error('Failed to attach MCP instructions to Microsoft Copilot:', error);
      return false;
    }
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
    const adapter = new Microsoft365CopilotBridgeAdapter();

    // The parent adapter owns the capability array. It is a readonly property,
    // but the array itself is intentionally mutable; advertise attachment support
    // before registration so useCurrentAdapter/MCPPopover enables the Attach action.
    if (!adapter.capabilities.includes('file-attachment')) {
      adapter.capabilities.push('file-attachment');
    }

    await pluginRegistry.register(adapter, {
      id: 'm365-copilot-adapter',
      name: 'Microsoft Copilot Adapter',
      description: 'Adapter for Microsoft Copilot Chat on copilot.cloud.microsoft and m365.cloud.microsoft',
      version: '1.3.0',
      enabled: true,
      priority: 10,
      settings: {
        logLevel: 'info',
      },
    });
    logger.debug('Microsoft Copilot adapter registered');
  }

  await pluginRegistry.activatePlugin(pluginName);

  // Microsoft currently renders some code blocks through a virtualized artifact
  // viewer whose DOM shape is not stable enough for the generic CSS-selector
  // renderer. This protocol-text bridge is deliberately M365-only and feeds a
  // normalized copy of completed JSONL requests into the existing Run-card UI.
  initializeMicrosoftCopilotFunctionBridge();

  logger.debug(`Microsoft Copilot adapter activated for ${hostname}`);
}
