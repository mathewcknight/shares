// Netlify serverless function — proxies Yahoo Finance chart API to avoid CORS.
// Endpoint: /.netlify/functions/prices?ticker=NDQ.AX

exports.handler = async (event) => {
  const ticker = event.queryStringParameters?.ticker

  if (!ticker) {
    return {
      statusCode: 400,
      headers: corsHeaders(),
      body: JSON.stringify({ error: 'Missing required query parameter: ticker' }),
    }
  }

  const url =
    `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}` +
    `?interval=1mo&range=20y`

  let response
  try {
    response = await fetch(url, {
      headers: {
        // Yahoo requires a recognisable User-Agent from server-side requests
        'User-Agent': 'Mozilla/5.0 (compatible; portfolio-dashboard/1.0)',
      },
    })
  } catch (err) {
    return {
      statusCode: 502,
      headers: corsHeaders(),
      body: JSON.stringify({ error: 'Failed to reach Yahoo Finance', detail: err.message }),
    }
  }

  const body = await response.text()

  return {
    statusCode: response.status,
    headers: {
      ...corsHeaders(),
      'Content-Type': 'application/json',
    },
    body,
  }
}

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  }
}
