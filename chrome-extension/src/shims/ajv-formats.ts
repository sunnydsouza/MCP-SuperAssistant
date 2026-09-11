// CSP-safe stub for `ajv-formats`.
//
// Some dependencies import `ajv-formats` or deep subpaths such as
// `ajv-formats/dist/formats`. The real package expects Ajv internals/codegen,
// which can trigger Manifest V3 CSP failures in the extension service worker.

export type AjvLike = {
  addFormat?: (name: string, format: unknown) => unknown;
  addKeyword?: (keyword: string, definition?: unknown) => unknown;
};

export default function addFormats(ajv: AjvLike, _opts?: unknown): AjvLike {
  return ajv;
}

export const formats: Record<string, unknown> = {};
export const fullFormats: Record<string, unknown> = {};
export const fastFormats: Record<string, unknown> = {};
