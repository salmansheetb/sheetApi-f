/**
 * worker/src/router.js
 * Minimal URL router for Cloudflare Workers.
 * Supports :param segments and all HTTP methods.
 */

export class Router {
  constructor() {
    this.routes = [];
  }

  add(method, pattern, handler) {
    // Escape only literal dots, then convert :param → named capture group
    const regexStr =
      '^' +
      pattern
        .replace(/\./g, '\\.')
        .replace(/:([A-Za-z_][A-Za-z0-9_]*)/g, '(?<$1>[^/]+)')
      + '$';
    this.routes.push({ method, regex: new RegExp(regexStr), handler });
  }

  get(pattern, handler)    { this.add('GET',    pattern, handler); }
  post(pattern, handler)   { this.add('POST',   pattern, handler); }
  put(pattern, handler)    { this.add('PUT',    pattern, handler); }
  delete(pattern, handler) { this.add('DELETE', pattern, handler); }

  async handle(request, env, ctx) {
    const url      = new URL(request.url);
    const pathname = url.pathname;
    const method   = request.method.toUpperCase();

    for (const route of this.routes) {
      if (route.method !== method) continue;
      const match = pathname.match(route.regex);
      if (!match) continue;
      const params = match.groups ?? {};
      return route.handler(request, { params, env, ctx });
    }

    return jsonResponse({ success: false, error: 'Route not found.' }, 404);
  }
}

export function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

export function errorResponse(err) {
  const status = err.statusCode ?? 500;
  return jsonResponse(
    { success: false, error: err.message ?? 'Internal Server Error' },
    status
  );
}
