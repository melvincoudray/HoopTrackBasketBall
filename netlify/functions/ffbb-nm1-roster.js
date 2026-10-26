// Fonction Netlify — lit l'effectif d'une équipe sur le site officiel de la NM1
// (page "Composition de l'équipe") : nom, numéro, taille, poste, date de naissance et photo.
// Appel : /.netlify/functions/ffbb-nm1-roster?team=<nom de l'équipe tel qu'il apparaît dans le tableau NM1>
// Les moyennes de saison de chaque joueur (fiche joueur) sont jointes à l'effectif.
// Les photos sont récupérées ici (serveur à serveur) et renvoyées en data-URL, car le navigateur
// ne peut généralement pas les lire directement (CORS). Aucune dépendance externe.

const ROOT = "https://nm1.ffbb.com";
const UA = { "User-Agent": "HoopTrack/1.0 (roster sync)" };

function decodeEntities(s) {
  return String(s)
    .replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&#0?39;|&apos;/g, "'")
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCharCode(parseInt(n, 16)));
}
function norm(s) {
  return String(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[\s\-_'’.]/g, "");
}

// Liens des équipes présents dans le tableau des stats : [{ name, href }]
export function parseTeamLinks(html) {
  const out = [];
  const re = /<a\b[^>]*>/gi;
  let m;
  while ((m = re.exec(html))) {
    const tag = m[0];
    if (!/class="[^"]*\bteam\b[^"]*"/.test(tag)) continue;
    const href = (tag.match(/href="([^"]+)"/) || [])[1];
    const title = (tag.match(/title="([^"]*)"/) || [])[1];
    if (href && title && /\/equipe\//.test(href)) out.push({ name: decodeEntities(title).trim(), href });
  }
  return out;
}

// Effectif d'une page équipe : [{ ffbbId, name, jersey, height, position, birth, photoUrl }]
export function parseRoster(html) {
  const sec = html.split('class="team__roster"')[1];
  if (!sec) return [];
  const items = sec.split(/<li[^>]*team-entry__roster/).slice(1);
  const seen = new Set();
  const out = [];
  for (const it of items) {
    const href = (it.match(/href="[^"]*\/joueur\/(\d+)-[^"]*"/) || [])[1];
    if (!href || seen.has(href)) continue;
    seen.add(href);
    const name = decodeEntities((it.match(/<h3[^>]*>([\s\S]*?)<\/h3>/) || [])[1] || "").replace(/\s+/g, " ").trim();
    const jersey = (it.match(/Num[ée]ro de maillot\s*:\s*(\d+)/i) || [])[1] || "";
    const hm = it.match(/<p>\s*(\d)m(\d{2})\s*<\/p>/);
    const height = hm ? String(Number(hm[1]) * 100 + Number(hm[2])) : "";
    const position = (it.match(/Poste\s*:\s*([\d-]+)/i) || [])[1] || "";
    const birth = (it.match(/<p>\s*(\d{2}\/\d{2}\/\d{2,4})\s*<\/p>/) || [])[1] || "";
    const photoUrl = (it.match(/<img[^>]*src="([^"]+)"/) || [])[1] || "";
    if (name) out.push({ ffbbId: href, name, jersey, height, position, birth, photoUrl });
  }
  return out;
}

// Moyennes de la saison régulière d'un joueur (page fiche joueur) pour l'équipe et la saison données.
// Retourne { mj, min, pts, reb, ast, r2, t2, r3, t3, lfr, lft, ro, rd, stl, tov, blk, eff } ou null.
export function parsePlayerSeason(html, teamName, seasonLabel) {
  const part = html.split(/Moyennes(?:<\/span>)?\s+en saison r[ée]guli[èe]re/i)[1];
  if (!part) return null;
  const tbl = (part.match(/<table[\s\S]*?<\/table>/i) || [])[0];
  if (!tbl) return null;
  const num = v => { const n = parseFloat(String(v).replace(",", ".")); return isNaN(n) ? null : n; };
  const pair = v => { const [a, b] = String(v).split("-"); return [num(a), num(b)]; };
  for (const tr of tbl.split(/<tr[\s>]/i).slice(1)) {
    const cells = [...tr.matchAll(/<td[^>]*>([\s\S]*?)<\/td>/gi)].map(m => decodeEntities(m[1].replace(/<[^>]+>/g, " ")).replace(/\s+/g, " ").trim());
    if (cells.length < 19) continue;
    if (cells[0] !== seasonLabel || norm(cells[1]) !== norm(teamName) || !/^FRA-3$/i.test(cells[2])) continue;
    const [r2, t2] = pair(cells[8]), [r3, t3] = pair(cells[9]), [lfr, lft] = pair(cells[11]);
    return { mj: num(cells[3]), min: num(cells[4]), pts: num(cells[5]), reb: num(cells[6]), ast: num(cells[7]),
      r2, t2, r3, t3, lfr, lft, ro: num(cells[13]), rd: num(cells[14]), stl: num(cells[15]), tov: num(cells[16]), blk: num(cells[17]), eff: num(cells[18]) };
  }
  return null;
}

async function getPlayerStats(ffbbId, teamName, seasonLabel) {
  try {
    const r = await fetch(ROOT + "/saison-reguliere/joueur/" + ffbbId, { headers: UA });
    if (!r.ok) return null;
    return parsePlayerSeason(await r.text(), teamName, seasonLabel);
  } catch { return null; }
}

async function getPhoto(url) {
  if (!/^https:\/\/www\.proballers\.com\//.test(url || "")) return null;
  try {
    const r = await fetch(url, { headers: UA });
    if (!r.ok) return null;
    const type = r.headers.get("content-type") || "";
    if (!type.startsWith("image/")) return null;
    const buf = Buffer.from(await r.arrayBuffer());
    if (buf.length < 500) return null;
    // Sans vraie photo, le site renvoie une image par défaut de 150 px (les vraies font 300 px) :
    // on la repère par la largeur écrite dans l'en-tête PNG et on l'ignore.
    if (buf.readUInt32BE(0) === 0x89504e47 && buf.readUInt32BE(16) < 300) return null;
    return "data:" + type + ";base64," + buf.toString("base64");
  } catch { return null; }
}

export const handler = async (event) => {
  try {
    const team = ((event.queryStringParameters || {}).team || "").trim();
    if (!team) return { statusCode: 400, body: JSON.stringify({ error: "Missing ?team=" }) };
    const season = ((event.queryStringParameters || {}).season || "2026").replace(/\D/g, "") || "2026";
    const lres = await fetch(ROOT + "/saison-reguliere/leaders/equipes/" + season, { headers: UA });
    if (!lres.ok) return { statusCode: 502, body: JSON.stringify({ error: "NM1 site answered " + lres.status }) };
    const link = parseTeamLinks(await lres.text()).find(t => norm(t.name) === norm(team));
    if (!link) return { statusCode: 404, body: JSON.stringify({ error: "Team not found on the NM1 site: " + team }) };
    const tres = await fetch(link.href.startsWith("http") ? link.href : ROOT + link.href, { headers: UA });
    if (!tres.ok) return { statusCode: 502, body: JSON.stringify({ error: "NM1 team page answered " + tres.status }) };
    const players = parseRoster(await tres.text());
    const yy = Number(season);
    const seasonLabel = String(yy % 100).padStart(2, "0") + "-" + String((yy + 1) % 100).padStart(2, "0");
    await Promise.all(players.map(async p => {
      const [photo, stats] = await Promise.all([getPhoto(p.photoUrl), getPlayerStats(p.ffbbId, link.name, seasonLabel)]);
      p.photo = photo; p.stats = stats; delete p.photoUrl;
    }));
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json", "Cache-Control": "public, max-age=300, s-maxage=1800" },
      body: JSON.stringify({ team: link.name, fetchedAt: new Date().toISOString(), players }),
    };
  } catch (e) {
    return { statusCode: 502, body: JSON.stringify({ error: String(e.message || e) }) };
  }
};
