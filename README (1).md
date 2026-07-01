export default {
  async fetch(request, env) {
    const corsHeaders = {
      "Access-Control-Allow-Origin": env.ALLOWED_ORIGIN || "*",
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    };

    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: corsHeaders });
    }

    const owner = env.GITHUB_OWNER || "n3kvm";
    const repo = env.GITHUB_REPO || "Indicador-Documental";
    const workflow = env.GITHUB_WORKFLOW || "dashboard-github-pages.yml";
    const ref = env.GITHUB_REF || "main";

    if (request.method === "GET") {
      const response = await fetch(
        `https://api.github.com/repos/${owner}/${repo}/actions/workflows/${workflow}/runs?per_page=1`,
        {
          headers: {
            Authorization: `Bearer ${env.GITHUB_TOKEN}`,
            Accept: "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "indicador-documental-refresh",
          },
        }
      );

      const data = await response.json();

      if (!response.ok) {
        return json(
          {
            ok: false,
            github_status: response.status,
            detail: data,
          },
          response.status,
          corsHeaders
        );
      }

      const run = data.workflow_runs?.[0];

      if (!run) {
        return json(
          {
            ok: false,
            error: "No hay ejecuciones del workflow.",
          },
          404,
          corsHeaders
        );
      }

      return json(
        {
          ok: true,
          status: run.status,
          conclusion: run.conclusion,
          workflow_url: run.html_url,
          created_at: run.created_at,
          updated_at: run.updated_at,
        },
        200,
        corsHeaders
      );
    }

    if (request.method !== "POST") {
      return json({ ok: false, error: "Use GET or POST" }, 405, corsHeaders);
    }

    let payload = {};
    try {
      payload = await request.json();
    } catch {
      payload = {};
    }

    const anio = String(payload.anio || "").trim();
    const mes = String(payload.mes || "").trim().padStart(2, "0");

    if (
      !/^\d{4}$/.test(anio) ||
      !/^\d{2}$/.test(mes) ||
      Number(mes) < 1 ||
      Number(mes) > 12
    ) {
      return json(
        { ok: false, error: "Digite anio de 4 digitos y mes entre 01 y 12" },
        400,
        corsHeaders
      );
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

      return json(
        {
          ok: false,
          github_status: response.status,
          detail,
        },
        response.status,
        corsHeaders
      );
    }

    return json(
      {
        ok: true,
        message: "Dashboard refresh started",
        status: "queued",
      },
      202,
      corsHeaders
    );
  },
};

function json(payload, status, headers) {
  return new Response(JSON.stringify(payload, null, 2), {
    status,
    headers: {
      ...headers,
      "Content-Type": "application/json; charset=utf-8",
    },
  });
}
