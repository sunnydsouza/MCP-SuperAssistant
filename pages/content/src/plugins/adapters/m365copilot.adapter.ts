import { BaseAdapterPlugin } from './base.adapter';
import type { AdapterCapability, PluginContext } from '../plugin-types';
import { createLogger } from '@extension/shared/lib/logger';

const logger = createLogger('Microsoft365CopilotAdapter');

/**
 * Adapter for Microsoft Copilot Chat / Microsoft 365 Copilot.
 *
 * Microsoft currently exposes the experience on both copilot.cloud.microsoft
 * and m365.cloud.microsoft depending on entry point and tenant routing.
 * Generated class names change frequently, so this adapter deliberately
 * prefers stable accessibility attributes, semantic editor attributes and
 * geometry-based fallbacks around the composer.
 */
export class Microsoft365CopilotAdapter extends BaseAdapterPlugin {
  readonly name = 'Microsoft365CopilotAdapter';
  readonly version = '1.1.0';
  readonly hostnames = ['copilot.cloud.microsoft', 'm365.cloud.microsoft'];
  readonly capabilities: AdapterCapability[] = ['text-insertion', 'form-submission', 'dom-manipulation'];

  private readonly inputSelectors = [
    'textarea[placeholder*="message" i]',
    'textarea[aria-label*="message" i]',
    'textarea[data-testid*="chat" i]',
    '[contenteditable="true"][role="textbox"]',
    '[contenteditable="true"][data-lexical-editor="true"]',
    '[contenteditable="true"][aria-label*="message" i]',
    '[contenteditable="true"][aria-label*="prompt" i]',
  ];

  private readonly sendSelectors = [
    'button[aria-label*="send" i]',
    'button[title*="send" i]',
    'button[data-testid*="send" i]',
    'button[aria-label*="submit" i]',
    'button[title*="submit" i]',
    'button[data-testid*="submit" i]',
    'button[type="submit"]',
  ];

  private mcpPopoverContainer: HTMLElement | null = null;
  private mcpPopoverRoot: { unmount?: () => void } | null = null;
  private mutationObserver: MutationObserver | null = null;
  private injectionTimer: ReturnType<typeof setTimeout> | null = null;

  async initialize(context: PluginContext): Promise<void> {
    await super.initialize(context);
    this.context.logger.debug('Microsoft Copilot adapter initialized');
  }

  async activate(): Promise<void> {
    if (this.currentStatus === 'active') return;
    await super.activate();
    this.context.logger.debug('Microsoft Copilot adapter activated');

    this.setupComposerIntegration();
  }

  async deactivate(): Promise<void> {
    if (this.currentStatus === 'inactive' || this.currentStatus === 'disabled') return;
    this.cleanupComposerIntegration();
    await super.deactivate();
    this.context.logger.debug('Microsoft Copilot adapter deactivated');
  }

  async cleanup(): Promise<void> {
    this.cleanupComposerIntegration();
    await super.cleanup();
    this.context.logger.debug('Microsoft Copilot adapter cleaned up');
  }

  isSupported(): boolean {
    const hostname = window.location.hostname.toLowerCase();
    return (
      hostname === 'copilot.cloud.microsoft' ||
      hostname === 'm365.cloud.microsoft' ||
      hostname.endsWith('.copilot.cloud.microsoft') ||
      hostname.endsWith('.m365.cloud.microsoft')
    );
  }

  async insertText(text: string, options?: { targetElement?: HTMLElement }): Promise<boolean> {
    const input = options?.targetElement ?? this.findPromptInput();
    if (!input) {
      this.emitFailed('insertText', 'Microsoft Copilot prompt editor was not found');
      return false;
    }

    try {
      input.focus();

      const existing = this.readEditorValue(input);
      const separator = existing.length > 0 && !existing.endsWith('\n') ? '\n' : '';
      const appendedText = `${separator}${text}`;
      const nextValue = `${existing}${appendedText}`;

      if (input instanceof HTMLTextAreaElement || input instanceof HTMLInputElement) {
        this.setNativeInputValue(input, nextValue);
      } else if (input.isContentEditable) {
        this.appendToContentEditable(input, appendedText, nextValue);
      } else {
        this.emitFailed('insertText', `Unsupported prompt editor element: ${input.tagName}`);
        return false;
      }

      this.emitCompleted('insertText', { textLength: text.length }, { success: true, editor: input.tagName });
      return true;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger.error('Failed to insert text into Microsoft Copilot:', error);
      this.emitFailed('insertText', message);
      return false;
    }
  }

  async submitForm(options?: { formElement?: HTMLFormElement }): Promise<boolean> {
    const input = this.findPromptInput();

    try {
      const sendButton = this.findBestSendButton(input);
      if (sendButton) {
        sendButton.click();
        this.emitCompleted('submitForm', {}, { success: true, method: 'send-button' });
        return true;
      }

      const enclosingForm = input?.closest('form') as HTMLFormElement | null;
      const form = options?.formElement ?? enclosingForm;
      if (form) {
        form.requestSubmit();
        this.emitCompleted('submitForm', {}, { success: true, method: 'requestSubmit' });
        return true;
      }

      this.emitFailed('submitForm', 'Microsoft Copilot Send button or form was not found');
      return false;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger.error('Failed to submit Microsoft Copilot prompt:', error);
      this.emitFailed('submitForm', message);
      return false;
    }
  }

  /**
   * The ChatGPT and GitHub Copilot adapters inject MCPPopover into the host
   * composer. Do the same for Microsoft Copilot, but avoid depending on
   * generated Fluent/Griffel class names.
   */
  private setupComposerIntegration(): void {
    this.schedulePopoverInjection(100);

    if (this.mutationObserver) return;

    this.mutationObserver = new MutationObserver(() => {
      const existing = document.getElementById('mcp-m365-popover-container');
      if (!existing || !existing.isConnected) {
        this.schedulePopoverInjection(250);
      }
    });

    this.mutationObserver.observe(document.body, {
      childList: true,
      subtree: true,
    });
  }

  private schedulePopoverInjection(delay: number): void {
    if (this.injectionTimer) return;

    this.injectionTimer = setTimeout(() => {
      this.injectionTimer = null;
      void this.injectMCPPopoverWithRetry();
    }, delay);
  }

  private async injectMCPPopoverWithRetry(maxAttempts = 12): Promise<void> {
    if (document.getElementById('mcp-m365-popover-container')) return;

    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      const insertionPoint = this.findButtonInsertionPoint();
      if (insertionPoint) {
        await this.injectMCPPopover(insertionPoint);
        return;
      }

      await new Promise(resolve => setTimeout(resolve, attempt < 5 ? 300 : 750));
    }

    this.context.logger.warn('Microsoft Copilot composer found no safe insertion point for MCP button');
  }

  private findButtonInsertionPoint(): { container: Element; insertAfter: Element | null } | null {
    const input = this.findPromptInput();
    if (!input) return null;

    const inputRect = input.getBoundingClientRect();
    const inputCenterY = inputRect.top + inputRect.height / 2;

    // Prefer semantically labelled Add/Attach buttons when Microsoft exposes
    // an accessible name for the plus button.
    const semanticSelectors = [
      'button[aria-label*="attach" i]',
      'button[title*="attach" i]',
      'button[aria-label*="add" i]',
      'button[title*="add" i]',
      'button[data-testid*="attach" i]',
      'button[data-testid*="add" i]',
    ];

    for (const selector of semanticSelectors) {
      for (const button of Array.from(document.querySelectorAll<HTMLButtonElement>(selector))) {
        if (!this.isClickable(button)) continue;
        const rect = button.getBoundingClientRect();
        const centerY = rect.top + rect.height / 2;
        if (Math.abs(centerY - inputCenterY) <= 100 && button.parentElement) {
          return { container: button.parentElement, insertAfter: button };
        }
      }
    }

    // Geometry fallback: walk up from the editor until we find a composer-like
    // container, then choose the closest visible button to the left of the editor.
    let composer: HTMLElement | null = input.parentElement;
    for (let depth = 0; composer && depth < 8; depth += 1, composer = composer.parentElement) {
      const buttons = Array.from(composer.querySelectorAll<HTMLButtonElement>('button')).filter(button =>
        this.isClickable(button),
      );
      if (buttons.length === 0) continue;

      const leftCandidates = buttons
        .map(button => ({ button, rect: button.getBoundingClientRect() }))
        .filter(({ rect }) => {
          const centerY = rect.top + rect.height / 2;
          return Math.abs(centerY - inputCenterY) <= 110 && rect.left <= inputRect.left + 80;
        })
        .sort((a, b) => Math.abs(b.rect.right - inputRect.left) - Math.abs(a.rect.right - inputRect.left));

      const plusLike = leftCandidates.find(({ button }) => {
        const label = `${button.getAttribute('aria-label') ?? ''} ${button.getAttribute('title') ?? ''} ${button.textContent ?? ''}`.trim();
        return /(^|\s)(add|attach|plus)(\s|$)/i.test(label) || button.textContent?.trim() === '+';
      });

      const candidate = plusLike?.button ?? leftCandidates[leftCandidates.length - 1]?.button;
      if (candidate?.parentElement) {
        return { container: candidate.parentElement, insertAfter: candidate };
      }
    }

    // Last resort: place the MCP control immediately before the prompt editor.
    if (input.parentElement) {
      return { container: input.parentElement, insertAfter: null };
    }

    return null;
  }

  private async injectMCPPopover(insertionPoint: { container: Element; insertAfter: Element | null }): Promise<void> {
    if (document.getElementById('mcp-m365-popover-container')) return;

    const reactContainer = document.createElement('div');
    reactContainer.id = 'mcp-m365-popover-container';
    reactContainer.setAttribute('data-mcp-superassistant', 'm365-composer');
    reactContainer.style.display = 'inline-flex';
    reactContainer.style.alignItems = 'center';
    reactContainer.style.flex = '0 0 auto';
    reactContainer.style.margin = '0 4px';
    reactContainer.style.position = 'relative';
    reactContainer.style.zIndex = '2';

    const { container, insertAfter } = insertionPoint;
    if (insertAfter && insertAfter.parentNode === container) {
      container.insertBefore(reactContainer, insertAfter.nextSibling);
    } else if (container.firstChild) {
      container.insertBefore(reactContainer, container.firstChild);
    } else {
      container.appendChild(reactContainer);
    }

    this.mcpPopoverContainer = reactContainer;

    try {
      const React = await import('react');
      const ReactDOM = await import('react-dom/client');
      const { MCPPopover } = await import('../../components/mcpPopover/mcpPopover');

      if (!reactContainer.isConnected) return;

      const toggleStateManager = this.createToggleStateManager();
      const root = ReactDOM.createRoot(reactContainer);
      this.mcpPopoverRoot = root;

      root.render(
        React.createElement(MCPPopover, {
          toggleStateManager,
          adapterName: this.name,
        }),
      );

      this.context.logger.debug('MCP popover injected into Microsoft Copilot composer');
    } catch (error) {
      reactContainer.remove();
      this.mcpPopoverContainer = null;
      this.mcpPopoverRoot = null;
      this.context.logger.error('Failed to render Microsoft Copilot MCP popover:', error);
    }
  }

  private createToggleStateManager() {
    const context = this.context;

    const stateManager = {
      getState: () => {
        const uiState = context.stores.ui;
        const preferences = uiState?.preferences ?? {};
        return {
          mcpEnabled: uiState?.mcpEnabled ?? true,
          autoInsert: preferences.autoInsert ?? false,
          autoSubmit: preferences.autoSubmit ?? false,
          autoExecute: preferences.autoExecute ?? false,
        };
      },
      setMCPEnabled: (enabled: boolean) => {
        const uiState = context.stores.ui;
        if (uiState?.setMCPEnabled) uiState.setMCPEnabled(enabled, 'm365-mcp-popover');
        stateManager.updateUI();
      },
      setAutoInsert: (enabled: boolean) => {
        context.stores.ui?.updatePreferences?.({ autoInsert: enabled });
        stateManager.updateUI();
      },
      setAutoSubmit: (enabled: boolean) => {
        context.stores.ui?.updatePreferences?.({ autoSubmit: enabled });
        stateManager.updateUI();
      },
      setAutoExecute: (enabled: boolean) => {
        context.stores.ui?.updatePreferences?.({ autoExecute: enabled });
        stateManager.updateUI();
      },
      updateUI: () => {
        const container = document.getElementById('mcp-m365-popover-container');
        if (!container) return;
        container.dispatchEvent(
          new CustomEvent('mcp-toggle-state-updated', {
            detail: stateManager.getState(),
          }),
        );
      },
    };

    return stateManager;
  }

  private cleanupComposerIntegration(): void {
    if (this.injectionTimer) {
      clearTimeout(this.injectionTimer);
      this.injectionTimer = null;
    }

    if (this.mutationObserver) {
      this.mutationObserver.disconnect();
      this.mutationObserver = null;
    }

    try {
      this.mcpPopoverRoot?.unmount?.();
    } catch (error) {
      this.context.logger.warn('Failed to unmount Microsoft Copilot MCP popover:', error);
    }
    this.mcpPopoverRoot = null;

    const container = this.mcpPopoverContainer ?? document.getElementById('mcp-m365-popover-container');
    container?.remove();
    this.mcpPopoverContainer = null;
  }

  private findPromptInput(): HTMLElement | null {
    for (const selector of this.inputSelectors) {
      const candidates = Array.from(document.querySelectorAll<HTMLElement>(selector));
      for (let index = candidates.length - 1; index >= 0; index -= 1) {
        const candidate = candidates[index];
        if (candidate && this.isVisible(candidate) && !this.isDisabled(candidate)) {
          return candidate;
        }
      }
    }
    return null;
  }

  private findBestSendButton(input: HTMLElement | null): HTMLButtonElement | null {
    const candidates: HTMLButtonElement[] = [];

    for (const selector of this.sendSelectors) {
      for (const button of Array.from(document.querySelectorAll<HTMLButtonElement>(selector))) {
        if (!candidates.includes(button) && this.isClickable(button)) candidates.push(button);
      }
    }

    for (const button of Array.from(document.querySelectorAll<HTMLButtonElement>('button'))) {
      if (candidates.includes(button) || !this.isClickable(button)) continue;
      const label = `${button.getAttribute('aria-label') ?? ''} ${button.getAttribute('title') ?? ''} ${button.textContent ?? ''}`.trim();
      if (/\b(send|submit)\b/i.test(label)) candidates.push(button);
    }

    if (candidates.length === 0) return null;
    if (!input) return candidates[candidates.length - 1] ?? null;

    const inputRect = input.getBoundingClientRect();
    const inputCenterY = inputRect.top + inputRect.height / 2;

    const nearby = candidates
      .map(button => {
        const rect = button.getBoundingClientRect();
        const centerY = rect.top + rect.height / 2;
        const verticalDistance = Math.abs(centerY - inputCenterY);
        const horizontalDistance = Math.abs(rect.left - inputRect.right);
        return { button, score: verticalDistance * 5 + horizontalDistance };
      })
      .filter(({ button }) => {
        const rect = button.getBoundingClientRect();
        return Math.abs(rect.top + rect.height / 2 - inputCenterY) <= 160;
      })
      .sort((a, b) => a.score - b.score);

    return nearby[0]?.button ?? candidates[candidates.length - 1] ?? null;
  }

  private readEditorValue(element: HTMLElement): string {
    if (element instanceof HTMLTextAreaElement || element instanceof HTMLInputElement) {
      return element.value ?? '';
    }
    return element.innerText ?? element.textContent ?? '';
  }

  private setNativeInputValue(element: HTMLTextAreaElement | HTMLInputElement, value: string): void {
    const prototype = element instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(prototype, 'value')?.set;

    if (setter) setter.call(element, value);
    else element.value = value;

    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: value }));
    element.dispatchEvent(new Event('change', { bubbles: true }));
  }

  private appendToContentEditable(element: HTMLElement, appendedText: string, fallbackValue: string): void {
    const range = document.createRange();
    range.selectNodeContents(element);
    range.collapse(false);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);

    let inserted = false;
    try {
      inserted = document.execCommand('insertText', false, appendedText);
    } catch {
      inserted = false;
    }

    if (!inserted) {
      element.textContent = fallbackValue;
      const fallbackRange = document.createRange();
      fallbackRange.selectNodeContents(element);
      fallbackRange.collapse(false);
      selection?.removeAllRanges();
      selection?.addRange(fallbackRange);
      element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: appendedText }));
    }

    element.dispatchEvent(new Event('change', { bubbles: true }));
  }

  private isVisible(element: HTMLElement): boolean {
    const style = window.getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    return style.display !== 'none' && style.visibility !== 'hidden' && Number(style.opacity) !== 0 && rect.width > 0 && rect.height > 0;
  }

  private isDisabled(element: HTMLElement): boolean {
    return (
      element.getAttribute('aria-disabled') === 'true' ||
      (element instanceof HTMLInputElement && element.disabled) ||
      (element instanceof HTMLTextAreaElement && element.disabled)
    );
  }

  private isClickable(button: HTMLButtonElement): boolean {
    return !button.disabled && button.getAttribute('aria-disabled') !== 'true' && this.isVisible(button);
  }

  private emitCompleted(toolName: string, parameters: Record<string, unknown>, result: Record<string, unknown>): void {
    this.context.eventBus.emit('tool:execution-completed', {
      execution: {
        id: this.generateCallId(),
        toolName,
        parameters,
        result,
        timestamp: Date.now(),
        status: 'success',
      },
    });
  }

  private emitFailed(toolName: string, error: string): void {
    this.context.logger.warn(`[Microsoft365CopilotAdapter] ${toolName} failed: ${error}`);
    this.context.eventBus.emit('tool:execution-failed', {
      toolName,
      error,
      callId: this.generateCallId(),
    });
  }

  private generateCallId(): string {
    return `m365-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
  }
}
