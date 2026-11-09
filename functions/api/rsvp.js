
export async function onRequestPost({ request, env }) {
  const headers = {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
  };

  try {
    const payload = await request.json();

    if (!payload) {
      return new Response(JSON.stringify({
        status: "error",
        message: "Invalid RSVP request."
      }), { status: 400, headers });
    }

    const response = await fetch(env.APPS_SCRIPT_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload),
      redirect: "follow"
    });

    const data = await response.text();
    JSON.parse(data);

    return new Response(data, {
      status: response.ok ? 200 : 502,
      headers
    });

  } catch (error) {
    return new Response(JSON.stringify({
      status: "error",
      message: "Unable to submit your RSVP. Please try again."
    }), { status: 502, headers });
  }
}
