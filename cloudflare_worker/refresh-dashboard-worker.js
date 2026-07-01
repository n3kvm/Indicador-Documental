export default {
  async fetch(request, env) {
    const corsHeaders = {
      "Access-Control-Allow-Origin": env.ALLOWED_ORIGIN || "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    };

    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: corsHeaders });
    }

    if (request.method !== "POST") {
      return json({ ok: false, error: "Use POST" }, 405, corsHeaders);
    }

    const owner = env.GITHUB_OWNER || "n3kvm";
    const repo = env.GITHUB_REPO || "Indicador-Documental";
    const workflow = env.GITHUB_WORKFLOW || "dashboard-github-pages.yml";
    const ref = env.GITHUB_REF || "main";
    let payload = {};
    try {
      payload = await request.json();
    } catch {
      payload = {};
    }

    const anio = String(payload.anio || "").trim();
    const mes = String(payload.mes || "").trim().padStart(2, "0");

    if (!/^\d{4}$/.test(anio) || !/^\d{2}$/.test(mes) || Number(mes) < 1 || Number(mes) > 12) {
      return json({ ok: false, error: "Digite anio de 4 digitos y mes entre 01 y 12" }, 400, corsHeaders);
    }

    const response = await fetch(
      `https://api.github.com/repos/${owner}/${repo}/actions/workflows/${workflow}/dispatches`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${env.GITHUB_TOKEN}`,
          Accept: "application/vnd.github+json",
          "X-GitHub-Api-Version": "2022-11-28",
          "User-Agent": "indicador-documental-refresh",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          ref,
          inputs: {
            anio,
            mes,
          },
        }),
      }
    );

    if (!response.ok) {
      const detail = await response.text();
      return json({ ok: false, status: response.status, detail }, 500, corsHeaders);
    }

    return json({ ok: true, message: "Dashboard refresh started" }, 202, corsHeaders);
  },
};

function json(payload, status, headers) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      ...headers,
      "Content-Type": "application/json; charset=utf-8",
    },
  });
}
