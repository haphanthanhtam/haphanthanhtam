export async function onRequest(context) {
  const request = context.request;
  const url = new URL(request.url);
  const target = url.searchParams.get("url");

  if (request.method === "OPTIONS") {
    return new Response(null, {
      status: 204,
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, HEAD, OPTIONS",
        "Access-Control-Allow-Headers": "*",
        "Access-Control-Max-Age": "86400"
      }
    });
  }

  if (!target) {
    return new Response(JSON.stringify({ error: "Missing url parameter" }), {
      status: 400,
      headers: {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*"
      }
    });
  }

  try {
    const resp = await fetch(target, {
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
      }
    });

    const headers = new Headers();
    headers.set("Access-Control-Allow-Origin", "*");
    headers.set("Access-Control-Allow-Methods", "GET, HEAD, OPTIONS");

    const contentType = resp.headers.get("content-type");
    if (contentType) headers.set("Content-Type", contentType);
    const contentLength = resp.headers.get("content-length");
    if (contentLength) headers.set("Content-Length", contentLength);
    headers.set("Cache-Control", "public, max-age=86400");

    return new Response(resp.body, {
      status: resp.status,
      headers: headers
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), {
      status: 502,
      headers: {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*"
      }
    });
  }
}
