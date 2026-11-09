// Trigger fresh Cloudflare deployment
export async function onRequestGet({ env }) {
  const headers = {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
  };

  let stage = "Starting";

  try {
    stage = "Checking Apps Script URL";

    if (!env.APPS_SCRIPT_URL) {
      throw new Error("APPS_SCRIPT_URL secret is missing.");
    }

    const url = new URL(env.APPS_SCRIPT_URL);
    url.searchParams.set("action", "getPublicConfig");

    stage = "Connecting to Google Apps Script";

    const response = await fetch(url.toString(), {
      method: "GET",
      redirect: "follow"
    });

    stage = "Reading Google Apps Script response";

    const data = await response.text();

    if (!response.ok) {
      throw new Error(`Google returned HTTP ${response.status}.`);
    }

    stage = "Checking JSON response";

    JSON.parse(data);

    return new Response(data, {
      status: 200,
      headers
    });

  } catch (error) {
    return new Response(JSON.stringify({
      status: "error",
      stage,
      message: error.message || String(error)
    }), {
      status: 502,
      headers
    });
  }
}
