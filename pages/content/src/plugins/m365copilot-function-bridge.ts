import { createLogger } from '@extension/shared/lib/logger';
import { renderFunctionCall } from '../render_prescript/src/renderer/functionBlock';
import { extractJSONObjects } from '../render_prescript/src/parser/jsonFunctionParser';

const logger = createLogger('Microsoft365CopilotFunctionBridge');

const FUNCTION_START = 'function_call_start';
const FUNCTION_END = 'function_call_end';
const BRIDGE_HOST_ATTRIBUTE = 'data-mcp-m365-function-bridge';
const SOURCE_ATTRIBUTE = 'data-mcp-m365-function-source';
const SCAN_INTERVAL_MS = 750;

const CODE_SOURCE_SELECTORS = [
  'pre',
  'code',
  '[role="code"]',
  '[data-testid*="code" i]',
  '[class*="code-block" i]',
  '[class*="codeBlock" i]',
  '.monaco-editor',
  '.monaco-editor .view-lines',
  '.view-lines',
].join(',');

let bridgeObserver: MutationObserver | null = null;
let bridgePollTimer: ReturnType<typeof setInterval> | null = null;
let scanTimer: ReturnType<typeof setTimeout> | null = null;
let initialized = false;
let bridgeSequence = 0;

// Track DOM instances rather than only protocol text. Copilot can legitimately
// emit the exact same function request more than once in one conversation.
const processedAnchors = new WeakSet<HTMLElement>();

type SearchRoot = Document | ShadowRoot;

function isMicrosoftCopilotHost(): boolean {
  const hostname = window.location.hostname.toLowerCase();
  return (
    hostname === 'copilot.cloud.microsoft' ||
    hostname === 'm365.cloud.microsoft' ||
    hostname.endsWith('.copilot.cloud.microsoft') ||
    hostname.endsWith('.m365.cloud.microsoft')
  );
}

function containsCompleteFunctionProtocol(text: string): boolean {
  return text.includes(FUNCTION_START) && text.includes(FUNCTION_END) && text.includes('"type"');
}

function normalizeFunctionProtocol(text: string): string | null {
  if (!containsCompleteFunctionProtocol(text)) return null;

  const jsonObjects = extractJSONObjects(text);
  if (jsonObjects.length === 0) return null;

  const protocolObjects: string[] = [];
  let insideFunction = false;

  for (const jsonObject of jsonObjects) {
    try {
      const parsed = JSON.parse(jsonObject) as { type?: string };
      if (parsed.type === FUNCTION_START) {
        insideFunction = true;
        protocolObjects.length = 0;
      }

      if (insideFunction) {
        protocolObjects.push(JSON.stringify(parsed));
      }

      if (insideFunction && parsed.type === FUNCTION_END) {
        break;
      }
    } catch {
      // extractJSONObjects already validated this object. Ignore a defensive
      // parse failure and continue looking for the protocol boundaries.
    }
  }

  if (protocolObjects.length === 0) return null;

  const normalized = protocolObjects.join('\n');
  return containsCompleteFunctionProtocol(normalized) ? normalized : null;
}

function signatureFor(content: string): string {
  let hash = 5381;
  for (let index = 0; index < content.length; index += 1) {
    hash = (hash * 33) ^ content.charCodeAt(index);
  }
  return (hash >>> 0).toString(36);
}

function findSmallestProtocolAncestor(start: HTMLElement, boundary: SearchRoot): HTMLElement | null {
  let current: HTMLElement | null = start;

  for (let depth = 0; current && depth < 14; depth += 1) {
    if (current.closest('.function-block') || current.hasAttribute(BRIDGE_HOST_ATTRIBUTE)) return null;

    const text = current.textContent ?? '';
    if (containsCompleteFunctionProtocol(text)) {
      return current;
    }

    const parent = current.parentElement;
    if (!parent || parent === document.body || parent === document.documentElement) break;

    if (boundary instanceof ShadowRoot && parent.getRootNode() !== boundary) break;
    current = parent;
  }

  return null;
}

function collectProtocolSources(root: SearchRoot): HTMLElement[] {
  const sources = new Set<HTMLElement>();

  for (const element of Array.from(root.querySelectorAll<HTMLElement>(CODE_SOURCE_SELECTORS))) {
    if (element.closest('.function-block') || element.closest(`[${BRIDGE_HOST_ATTRIBUTE}]`)) continue;
    if (containsCompleteFunctionProtocol(element.textContent ?? '')) {
      sources.add(element);
    }
  }

  // Microsoft changes the implementation of its code/artifact viewer regularly.
  // If none of the semantic/code selectors matched, locate the text itself and
  // climb only as far as the smallest common ancestor containing start + end.
  const rootText = root instanceof Document ? root.body?.textContent ?? '' : root.textContent ?? '';
  if (sources.size === 0 && containsCompleteFunctionProtocol(rootText)) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let node = walker.nextNode();

    while (node) {
      const text = node.textContent ?? '';
      if (text.includes(FUNCTION_START)) {
        const parent = node.parentElement;
        if (parent) {
          const candidate = findSmallestProtocolAncestor(parent, root);
          if (candidate) sources.add(candidate);
        }
      }
      node = walker.nextNode();
    }
  }

  return Array.from(sources);
}

function findOuterCodeAnchor(source: HTMLElement): HTMLElement {
  const semanticAnchor = source.closest<HTMLElement>(
    'pre, [data-testid*="code" i], [class*="code-block" i], [class*="codeBlock" i], .monaco-editor',
  );
  return semanticAnchor ?? source;
}

function mountSyntheticFunctionBlock(anchor: HTMLElement, normalizedProtocol: string): boolean {
  if (processedAnchors.has(anchor) || anchor.hasAttribute(SOURCE_ATTRIBUTE)) return false;

  const signature = signatureFor(normalizedProtocol);
  const bridgeId = `${signature}-${++bridgeSequence}`;

  const bridgeHost = document.createElement('div');
  bridgeHost.setAttribute(BRIDGE_HOST_ATTRIBUTE, bridgeId);
  bridgeHost.style.display = 'block';
  bridgeHost.style.width = '100%';
  bridgeHost.style.margin = '8px 0';

  const syntheticPre = document.createElement('pre');
  syntheticPre.textContent = normalizedProtocol;
  syntheticPre.setAttribute('data-mcp-m365-normalized-protocol', 'true');
  bridgeHost.appendChild(syntheticPre);

  const root = anchor.getRootNode();

  if (root instanceof ShadowRoot) {
    const shadowHost = root.host as HTMLElement;
    if (shadowHost.parentElement) {
      shadowHost.insertAdjacentElement('afterend', bridgeHost);
    } else {
      document.body.appendChild(bridgeHost);
    }
  } else if (anchor.parentElement) {
    anchor.insertAdjacentElement('afterend', bridgeHost);
  } else {
    document.body.appendChild(bridgeHost);
  }

  const rendered = renderFunctionCall(syntheticPre, { current: false });
  if (!rendered) {
    bridgeHost.remove();
    logger.warn('M365 protocol bridge found JSONL, but the standard function renderer rejected the normalized block');
    return false;
  }

  anchor.setAttribute(SOURCE_ATTRIBUTE, bridgeId);
  processedAnchors.add(anchor);
  logger.info(`Rendered Microsoft Copilot MCP function request through fallback bridge (${bridgeId})`);
  return true;
}

function scanRoot(root: SearchRoot): number {
  let rendered = 0;
  const candidates = new Map<HTMLElement, string>();

  for (const source of collectProtocolSources(root)) {
    const anchor = findOuterCodeAnchor(source);
    if (processedAnchors.has(anchor) || anchor.hasAttribute(SOURCE_ATTRIBUTE)) continue;

    const normalized = normalizeFunctionProtocol(source.textContent ?? '');
    if (!normalized) continue;

    // Nested <pre>/<code>/Monaco selectors commonly point at the same code
    // artifact. Key by the outer anchor so one Copilot response gets one Run card.
    if (!candidates.has(anchor)) candidates.set(anchor, normalized);
  }

  for (const [anchor, normalized] of candidates) {
    if (mountSyntheticFunctionBlock(anchor, normalized)) rendered += 1;
  }

  return rendered;
}

function scanOpenShadowRoots(): number {
  let rendered = 0;
  const visited = new Set<ShadowRoot>();

  const scanElementTree = (root: Document | ShadowRoot): void => {
    for (const element of Array.from(root.querySelectorAll<HTMLElement>('*'))) {
      const shadowRoot = element.shadowRoot;
      if (!shadowRoot || visited.has(shadowRoot)) continue;

      visited.add(shadowRoot);
      rendered += scanRoot(shadowRoot);
      scanElementTree(shadowRoot);
    }
  };

  scanElementTree(document);
  return rendered;
}

function scanSameOriginFrames(): number {
  let rendered = 0;

  for (const iframe of Array.from(document.querySelectorAll<HTMLIFrameElement>('iframe'))) {
    try {
      const frameDocument = iframe.contentDocument;
      if (!frameDocument?.body) continue;

      for (const source of collectProtocolSources(frameDocument)) {
        const frameAnchor = findOuterCodeAnchor(source);
        if (processedAnchors.has(frameAnchor) || frameAnchor.hasAttribute(SOURCE_ATTRIBUTE)) continue;

        const normalized = normalizeFunctionProtocol(source.textContent ?? '');
        if (!normalized) continue;

        // The standard renderer is bound to the top-level content script's
        // document. Proxy only the normalized protocol back beside the iframe.
        const proxyAnchor = document.createElement('span');
        proxyAnchor.style.display = 'none';
        iframe.insertAdjacentElement('afterend', proxyAnchor);
        const didRender = mountSyntheticFunctionBlock(proxyAnchor, normalized);
        proxyAnchor.remove();

        if (didRender) {
          frameAnchor.setAttribute(SOURCE_ATTRIBUTE, signatureFor(normalized));
          processedAnchors.add(frameAnchor);
          rendered += 1;
        }
      }
    } catch {
      // Cross-origin frame. The top-level content script cannot inspect it.
    }
  }

  return rendered;
}

function scanForMicrosoftFunctionCalls(): void {
  if (!isMicrosoftCopilotHost()) return;

  const renderedInDocument = scanRoot(document);
  if (renderedInDocument > 0) return;

  // These are intentionally fallbacks. A normal Microsoft code block should be
  // found in the top-level document without traversing every shadow host/frame.
  const renderedInShadowRoots = scanOpenShadowRoots();
  if (renderedInShadowRoots > 0) return;

  scanSameOriginFrames();
}

function scheduleScan(delay = 80): void {
  if (scanTimer) return;
  scanTimer = setTimeout(() => {
    scanTimer = null;
    scanForMicrosoftFunctionCalls();
  }, delay);
}

export function initializeMicrosoftCopilotFunctionBridge(): void {
  if (initialized || !isMicrosoftCopilotHost()) return;
  initialized = true;

  scheduleScan(0);

  bridgeObserver = new MutationObserver(mutations => {
    for (const mutation of mutations) {
      const text =
        mutation.type === 'characterData'
          ? mutation.target.textContent ?? ''
          : Array.from(mutation.addedNodes)
              .map(node => node.textContent ?? '')
              .join('');

      if (text.includes(FUNCTION_START) || text.includes(FUNCTION_END)) {
        scheduleScan();
        return;
      }
    }
  });

  bridgeObserver.observe(document.body, {
    childList: true,
    subtree: true,
    characterData: true,
  });

  // Polling is a safety net for virtualized Microsoft response/code viewers that
  // replace subtrees in ways that can make mutation timing racy.
  bridgePollTimer = setInterval(scanForMicrosoftFunctionCalls, SCAN_INTERVAL_MS);

  window.addEventListener(
    'unload',
    () => {
      bridgeObserver?.disconnect();
      bridgeObserver = null;
      if (bridgePollTimer) clearInterval(bridgePollTimer);
      bridgePollTimer = null;
      if (scanTimer) clearTimeout(scanTimer);
      scanTimer = null;
      initialized = false;
    },
    { once: true },
  );

  logger.info('Microsoft Copilot JSONL function fallback bridge initialized');
}
