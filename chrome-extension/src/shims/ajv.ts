// CSP-safe Ajv shim for Chrome/Edge Manifest V3.
//
// Ajv normally compiles schemas with `new Function()`, which is blocked by
// Manifest V3 extension CSP. The MCP SDK uses JSON-schema validation while
// discovering primitives/tools, so allowing real Ajv codegen into the service
// worker can leave the extension "connected" but with an empty tools list.
//
// This shim intentionally makes validation permissive inside the extension.
// Validation still occurs on the actual MCP server side where applicable.

export class Name {
  public str: string;

  constructor(s: unknown) {
    this.str = String(s);
  }

  toString() {
    return this.str;
  }
}

export default class Ajv {
  constructor(_options?: unknown) {}

  // Return `this` so common Ajv method chaining continues to work.
  addSchema(_schema: unknown, _key?: string) {
    return this;
  }

  addKeyword(_keyword: string, _definition?: unknown) {
    return this;
  }

  addFormat(_name: string, _format: unknown) {
    return this;
  }

  compile(_schema: unknown) {
    return (_data: unknown) => true;
  }

  validate(_schema: unknown, _data: unknown) {
    return true;
  }

  errors = null;
}

// Defensive exports for dependencies that import Ajv codegen symbols directly.
export const _ = () => {};
export const str = () => {};
export const nil = {};
export const CodeGen = {};
