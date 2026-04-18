/**
 * worker/src/index.js
 * Cloudflare Worker entry point.
 */

import { buildRouter } from './routes.js';
import { jsonResponse } from './router.js';

const router = buildRouter();

const CORS_HEADERS = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
};

function withCors(response) {
  const res = new Response(response.body, response);
  for (const [k, v] of Object.entries(CORS_HEADERS)) res.headers.set(k, v);
  return res;
}

export default {
  async fetch(request, env, ctx) {
    // CORS preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS_HEADERS });
    }

    try {
      const response = await router.handle(request, env, ctx);
      return withCors(response);
    } catch (err) {
      return withCors(
        jsonResponse({ success: false, error: err.message ?? 'Internal Server Error' }, 500)
      );
    }
  },
};
