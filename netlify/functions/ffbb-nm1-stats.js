// Fonction Netlify — lit le tableau "Statistiques équipes NM1" du site officiel de la NM1
// (https://nm1.ffbb.com/saison-reguliere/leaders/equipes/<saison>) et le renvoie en JSON à
// l'application, qui le traite avec son import "Classic" existant (mêmes colonnes : EQUIPE, MJ,
// PTS, EVA, 2R, 2T, 2%, 3R, 3T, 3%, LFR, LFT, LF%, RT, RO, RD, PD, IN, BP, CT, FTE, FPR).
// Pourquoi côté serveur : le navigateur ne peut généralement pas lire ce site directement
// (blocage CORS) ; ici la lecture se fait de serveur à serveur, sans dépendance externe.
// Le résultat est mis en cache 30 minutes côté Netlify pour ne pas solliciter le site à chaque
// ouverture de l'application.

const BASE = "https://nm1.ffbb.com/saison-reguliere/leaders/equipes/";

function decodeEntities(s) {
  return s
    .replace(/&nbsp;|&#160;/g, " ")
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&#0?39;|&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)));
}
function cellText(html) {
  return decodeEntities(html.replace(/<[^>]*>/g, " ")).replace(/[\s ]+/g, " ").trim();
}

// Extrait { headers, rows } du tableau #myTable. Exporté pour pouvoir être testé hors Netlify.
export function parseNm1Table(html) {
  const tableMatch = html.match(/<table[^>]*id=["']myTable["'][^>]*>([\s\S]*?)<\/table>/i)
    || html.match(/<table[^>]*>([\s\S]*?)<\/table>/i);
  if (!tableMatch) throw new Error("Table not found in the page");
  const tableHtml = tableMatch[1];
  const theadMatch = tableHtml.match(/<thead[^>]*>([\s\S]*?)<\/thead>/i);
  const headers = theadMatch ? [...theadMatch[1].matchAll(/<th[^>]*>([\s\S]*?)<\/th>/gi)].map(m => cellText(m[1])) : [];
  const bodyHtml = (tableHtml.match(/<tbody[^>]*>([\s\S]*?)<\/tbody>/i) || [null, tableHtml])[1];
  const rows = [];
  for (const tr of bodyHtml.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)) {
    const cells = [...tr[1].matchAll(/<td[^>]*>([\s\S]*?)<\/td>/gi)].map(m => cellText(m[1]));
    if (cells.length >= 5 && cells[0]) rows.push(cells);
  }
  if (!headers.length || !rows.length) throw new Error("Table is empty or has an unexpected structure");
  return { headers, rows };
}

// Points encaissés par match de chaque équipe, lus dans le classement officiel
// (https://nm1.ffbb.com/saison-reguliere/classement — un tableau par poule) : { "Nom équipe": moyenne }.
// Colonnes : Position, Equipe, Points, MJ, V, D, Pts marqués (total/moy), Pts encaissés (total/moy), Différence.
export function parseStandingsAllowed(html) {
  const out = {};
  for (const tr of html.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)) {
    const cells = [...tr[1].matchAll(/<td[^>]*>([\s\S]*?)<\/td>/gi)].map(m => cellText(m[1]));
    if (cells.length < 10) continue;
    const avg = parseFloat(String(cells[9]).replace(",", "."));
    if (cells[1] && !isNaN(avg)) out[cells[1]] = avg;
  }
  return out;
}

export const handler = async (event) => {
  const season = String((event.queryStringParameters && event.queryStringParameters.season) || "2026").replace(/[^0-9]/g, "") || "2026";
  try {
    const res = await fetch(BASE + season, { headers: { "User-Agent": "HoopTrack/1.0 (team stats sync)", "Accept": "text/html" } });
    if (!res.ok) return { statusCode: 502, body: JSON.stringify({ error: "NM1 site answered " + res.status }) };
    const { headers, rows } = parseNm1Table(await res.text());
    // Points encaissés (classement) : un bonus — s'il échoue, le reste de la synchro fonctionne quand même.
    let allowed = {};
    try {
      const cres = await fetch("https://nm1.ffbb.com/saison-reguliere/classement", { headers: { "User-Agent": "HoopTrack/1.0 (team stats sync)", "Accept": "text/html" } });
      if (cres.ok) allowed = parseStandingsAllowed(await cres.text());
    } catch { allowed = {}; }
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json", "Cache-Control": "public, max-age=300, s-maxage=1800" },
      body: JSON.stringify({ season, fetchedAt: new Date().toISOString(), headers, rows, allowed }),
    };
  } catch (e) {
    return { statusCode: 502, body: JSON.stringify({ error: String(e.message || e) }) };
  }
};
