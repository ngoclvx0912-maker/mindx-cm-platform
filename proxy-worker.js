// Cloudflare Worker — Proxy for Google Apps Script
// Deploy this to Cloudflare Workers (free tier: 100K requests/day)

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyJpFo_X0M_uvSPCStTOPVIAganyFRaN2AaCGO-Ukv911nlVS3me-C0jfIc8AiTSF1V/exec';

export default {
  async fetch(request) {
    const url = new URL(request.url);
    
    // CORS headers
    const corsHeaders = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    };

    // Handle preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders });
    }

    // Forward to Google Apps Script
    const targetUrl = APPS_SCRIPT_URL + url.search;
    
    const fetchOptions = {
      method: request.method,
      redirect: 'follow',
    };

    if (request.method === 'POST') {
      fetchOptions.body = await request.text();
      fetchOptions.headers = { 'Content-Type': 'text/plain' };
    }

    try {
      const response = await fetch(targetUrl, fetchOptions);
      const body = await response.text();
      
      return new Response(body, {
        status: response.status,
        headers: {
          'Content-Type': 'application/json',
          ...corsHeaders,
        },
      });
    } catch (err) {
      return new Response(JSON.stringify({ error: err.message }), {
        status: 500,
        headers: { 'Content-Type': 'application/json', ...corsHeaders },
      });
    }
  },
};
