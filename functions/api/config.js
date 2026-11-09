
export async function onRequestGet({ env }) {
  const headers = {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
  };

  try {
    const url = new URL(env.APPS_SCRIPT_URL);
    url.searchParams.set("action", "getPublicConfig");

    const response = await fetch(url.toString(), {
      method: "GET",
      redirect: "follow"
    });

    if (!response.ok) {
      throw new Error("Apps Script request failed");
    }

    const data = await response.text();
    JSON.parse(data);

    return new Response(data, { status: 200, headers });

  } catch (error) {
    return new Response(JSON.stringify({
      status: "error",
      message: "Could not load invitation settings."
    }), {
      status: 502,
      headers
    });
  }
}
