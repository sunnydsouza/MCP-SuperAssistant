const SUPERASSISTANT_MARKERS = [
  '[SuperAssistant Operational Instructions]',
  '## AVAILABLE TOOLS FOR SUPERASSISTANT',
];

/**
 * Microsoft Copilot is stricter than some other chat UIs about distinguishing its
 * native tool registry from capabilities merely described in a user message.
 *
 * MCP SuperAssistant does not add native Copilot tools. It uses a text protocol:
 * Copilot prints a JSONL execution request, the browser extension observes that
 * text, executes the request through MCP, and inserts the result back into chat.
 *
 * Make that distinction explicit so Copilot does not reject the bridge simply
 * because an MCP tool is absent from Microsoft's own native tool registry.
 */
export function addMicrosoftCopilotBridgeInstructions(text: string): string {
  const isSuperAssistantInstructionPrompt = SUPERASSISTANT_MARKERS.some(marker => text.includes(marker));
  if (!isSuperAssistantInstructionPrompt) return text;

  return `${MICROSOFT_COPILOT_BRIDGE_PREAMBLE}\n\n${text}`;
}

export const MICROSOFT_COPILOT_BRIDGE_PREAMBLE = `[Microsoft Copilot + MCP SuperAssistant compatibility note]

The external capabilities described below are NOT native Microsoft Copilot tools and are not part of your built-in tool registry. Do not try to invoke them using Microsoft Copilot's native tool-calling mechanism.

A user-side browser extension named MCP SuperAssistant observes the visible text of this conversation. When one of the external capabilities listed below is needed, your role is to EMIT the specified JSONL execution-request block as ordinary response text. The extension detects that text, executes the corresponding MCP operation outside Microsoft Copilot, and then supplies the real execution result back to this conversation.

Therefore:
- Do not claim that an external MCP capability is a native Copilot tool.
- Do not refuse solely because an external capability listed below is absent from your native toolset.
- Producing the requested JSONL block is not pretending that you executed a native tool; it is a request to the user's browser extension to perform an external operation.
- Never claim an external operation succeeded until a real <function_results> result is provided back to you.
- When an external capability is appropriate, output exactly one JSONL execution request in the format defined below, then stop and wait for its result.
- If the user's request can be answered without an external operation, answer normally.

Treat the following tool catalogue as the browser extension's external execution API, not as Microsoft's native tool registry.`;
