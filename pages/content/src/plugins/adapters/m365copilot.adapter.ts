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
 * prefers stable accessibility attributes and semantic editor attributes.
 */
export class Microsoft365CopilotAdapter extends BaseAdapterPlugin {
  readonly name = 'Microsoft365CopilotAdapter';
  readonly version = '1.0.1';
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

  async initialize(context: PluginContext): Promise<void> {
    await super.initialize(context);
    this.context.logger.debug('Microsoft Copilot adapter initialized');
  }

  async activate(): Promise<void> {
    if (this.currentStatus === 'active') return;
    await super.activate();
    this.context.logger.debug('Microsoft Copilot adapter activated');
  }

  async deactivate(): Promise<void> {
    if (this.currentStatus === 'inactive' || this.currentStatus === 'disabled') return;
    await super.deactivate();
    this.context.logger.debug('Microsoft Copilot adapter deactivated');
  }

  async cleanup(): Promise<void> {
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
