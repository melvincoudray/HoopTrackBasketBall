// Fonction Netlify — récupère les logos des clubs du tableau "Statistiques équipes NM1"
// (https://nm1.ffbb.com/saison-reguliere/leaders/equipes/<saison>). Les logos sont des SVG hébergés
// chez Proballers ; ils sont lus ici (serveur à serveur, le navigateur est bloqué par CORS) et
// renvoyés en data-URL : { teams: [{ name, logo }] }. Aucune dépendance externe.

const ROOT = "https://nm1.ffbb.com";
const UA = { "User-Agent": "HoopTrack/1.0 (logo sync)" };

function decodeEntities(s) {
  return String(s)
    .replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&#0?39;|&apos;/g, "'")
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">");
}

// [{ name, id }] : nom et identifiant de chaque club du tableau (identifiant = celui de l'adresse
// /equipe/<id>-<nom>, qui est aussi celui du logo).
export function parseTeamIds(html) {
  const out = [];
  const re = /<a\b[^>]*>/gi;
  let m;
  while ((m = re.exec(html))) {
    const tag = m[0];
    if (!/class="[^"]*\bteam\b[^"]*"/.test(tag)) continue;
    const id = (tag.match(/href="[^"]*\/equipe\/(\d+)-[^"]*"/) || [])[1];
    const title = (tag.match(/title="([^"]*)"/) || [])[1];
    if (id && title) out.push({ name: decodeEntities(title).trim(), id });
  }
  return out;
}

async function getLogo(id) {
  try {
    const r = await fetch("https://www.proballers.com/media/team/" + id + ".svg", { headers: UA });
    if (!r.ok) return null;
    const buf = Buffer.from(await r.arrayBuffer());
    const head = buf.toString("utf8", 0, 400).toLowerCase();
    if (buf.length < 100 || !head.includes("<svg")) return null;
    return "data:image/svg+xml;base64," + buf.toString("base64");
  } catch { return null; }
}

export const handler = async (event) => {
  try {
    const season = ((event.queryStringParameters || {}).season || "2026").replace(/\D/g, "") || "2026";
    const res = await fetch(ROOT + "/saison-reguliere/leaders/equipes/" + season, { headers: UA });
    if (!res.ok) return { statusCode: 502, body: JSON.stringify({ error: "NM1 site answered " + res.status }) };
    const teams = parseTeamIds(await res.text());
    if (!teams.length) return { statusCode: 502, body: JSON.stringify({ error: "No team found on the NM1 page." }) };
    const out = await Promise.all(teams.map(async t => ({ name: t.name, logo: await getLogo(t.id) })));
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json", "Cache-Control": "public, max-age=3600, s-maxage=86400" },
      body: JSON.stringify({ fetchedAt: new Date().toISOString(), teams: out }),
    };
  } catch (e) {
    return { statusCode: 502, body: JSON.stringify({ error: String(e.message || e) }) };
  }
};
