// Fonction Netlify — lit le classement officiel de la NM1, qui est scindé en plusieurs poules
// (https://nm1.ffbb.com/saison-reguliere/classement — un <h2>"Classement poule X" + un tableau
// par poule). Renvoyée telle quelle (par poule), pour que l'appli puisse :
//  - afficher le VRAI classement de la poule d'une équipe (pas seulement un classement local,
//    limité aux équipes déjà chargées dans Scouting) ;
//  - filtrer "le rang dans la ligue" (déjà calculé ailleurs à partir des équipes chargées dans
//    Scouting) à la seule poule de l'adversaire scouté, pour un "rang de poule" plus juste.
// Lecture seule, sans dépendance externe ; mise en cache 30 min côté Netlify (le classement ne
// bouge qu'après des journées de championnat).

const BASE = "https://nm1.ffbb.com/saison-reguliere/classement";
const UA = { "User-Agent": "HoopTrack/1.0 (poule standings sync)", "Accept": "text/html" };

function decodeEntities(s) {
  return String(s)
    .replace(/&nbsp;|&#160;/g, " ")
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&#0?39;|&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)));
}
function cellText(html) {
  return decodeEntities(String(html).replace(/<[^>]*>/g, " ")).replace(/[\s ]+/g, " ").trim();
}

// { "Poule A": [{ pos, name, pts, mj, w, l, pf, pfAvg, pa, paAvg, diff }], "Poule B": [...] }
export function parseClassementPoules(html) {
  const poules = {};
  const secRe = /<h2[^>]*class="generic__title"[^>]*>\s*<span>Classement<\/span>\s*poule\s*([^<]+?)\s*<\/h2>/gi;
  let m;
  const heads = [];
  while ((m = secRe.exec(html))) heads.push({ name: decodeEntities(m[1]).trim(), index: m.index });
  heads.forEach((h, i) => {
    const pouleName = "Poule " + h.name.toUpperCase();
    const chunk = html.slice(h.index, heads[i + 1] ? heads[i + 1].index : html.length);
    const tableMatch = chunk.match(/<table[^>]*class="table"[^>]*>([\s\S]*?)<\/table>/i);
    if (!tableMatch) return;
    const tableHtml = tableMatch[1];
    const bodyHtml = (tableHtml.match(/<tbody[^>]*>([\s\S]*?)<\/tbody>/i) || [null, tableHtml])[1];
    const rows = [];
    for (const tr of bodyHtml.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)) {
      const trHtml = tr[1];
      const cells = [...trHtml.matchAll(/<td[^>]*>([\s\S]*?)<\/td>/gi)].map(c => cellText(c[1]));
      if (cells.length < 11) continue;
      const teamMatch = trHtml.match(/<a[^>]*class="team"[^>]*title="([^"]*)"/i);
      const name = teamMatch ? decodeEntities(teamMatch[1]).trim() : cells[1];
      const num = v => { const n = parseFloat(String(v).replace(",", ".")); return isNaN(n) ? null : n; };
      rows.push({
        pos: num(cells[0]), name,
        pts: num(cells[2]), mj: num(cells[3]), w: num(cells[4]), l: num(cells[5]),
        pf: num(cells[6]), pfAvg: num(cells[7]), pa: num(cells[8]), paAvg: num(cells[9]), diff: num(cells[10]),
      });
    }
    if (rows.length) poules[pouleName] = rows.sort((a, b) => (a.pos ?? 999) - (b.pos ?? 999));
  });
  return poules;
}

export const handler = async () => {
  try {
    const res = await fetch(BASE, { headers: UA });
    if (!res.ok) return { statusCode: 502, body: JSON.stringify({ error: "NM1 site answered " + res.status }) };
    const poules = parseClassementPoules(await res.text());
    if (!Object.keys(poules).length) return { statusCode: 502, body: JSON.stringify({ error: "No poule tables found on the NM1 classement page." }) };
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json", "Cache-Control": "public, max-age=900, s-maxage=1800" },
      body: JSON.stringify({ fetchedAt: new Date().toISOString(), poules }),
    };
  } catch (e) {
    return { statusCode: 502, body: JSON.stringify({ error: String(e.message || e) }) };
  }
};
