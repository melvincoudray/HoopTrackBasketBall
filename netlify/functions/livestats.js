// Fonction Netlify — LiveStats (FIBA / Genius Sports) des matchs NM1, lues via le lien présent sur chaque match du calendrier FFBB.
//  • GET ?team=Vitré        -> { games: [{ live, home, away, dm }] }  (matchs JOUÉS de l'équipe ; "live" = identifiant LiveStats)
//  • GET ?live=<id>         -> { teams: {1:{name,starters,names}, 2:{...}}, ev: [[t, actionNumber, tno, shirt, type, sub, s1, s2, previousAction]], shots: [[tno, shirt, made, x, y]] }
//    t = secondes écoulées depuis le début du match (temps réglementaire 0-2400, prolongations ensuite), trié.
//    shots : un élément par tir à 2 ou 3 pts (hors lancers francs), pour la carte de tir par zone.
//    made = 1 si réussi / 0 si manqué ; x, y = position du tir en % du terrain ENTIER (confirmé le
//    30/10/2026 sur un vrai data.json : champ "shot" sous tm["1"]/tm["2"], pas dans "pbp" — x va de 0
//    à 100 sur la longueur (28 m), y de 0 à 100 sur la largeur (15 m) ; les deux équipes gardent leurs
//    coordonnées physiques réelles et changent donc de côté entre les périodes 1-2 et 3-4, comme au vrai jeu.
// Le navigateur ne peut pas lire data.json directement (pas de CORS) : cette fonction sert de relais.
const ROOT = "https://nm1.ffbb.com";
const LIVE = "https://fibalivestats.dcd.shared.geniussports.com";
const UA = { "User-Agent": "HoopTrack/1.0 (livestats sync)" };
const norm = s => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]/g, "");
const dec = s => String(s).replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&#0?39;|&apos;|&rsquo;/g, "'").replace(/&eacute;/g, "é").replace(/&egrave;/g, "è");
const headers = { "Content-Type": "application/json", "Cache-Control": "no-cache" };

async function getText(url) {
  const r = await fetch(url, { headers: UA });
  if (!r.ok) throw new Error("Source answered " + r.status);
  return r.text();
}
// Un bloc <li ...entry> par match : lien LiveStats (container href), "A vs B" (title), scores affichés si joué.
function parseGames(html) {
  const out = [];
  const re = /<li class="display-games__third-list__entry[^"]*">([\s\S]*?)<\/li>/g;
  let m;
  while ((m = re.exec(html))) {
    const li = m[1];
    const live = (li.match(/href="https?:\/\/fibalivestats[^"]*\/u\/FFBB\/(\d+)\/?"/) || [])[1];
    const title = (li.match(/<a class="header__game__link"[^>]*title="([^"]*)"/) || [])[1];
    if (!live || !title) continue;
    const scores = [...li.matchAll(/<span class="team__score[^"]*">\s*([^<]*?)\s*<\/span>/g)].map(x => x[1].trim());
    if (!/^\d+$/.test(scores[0] || "")) continue; // pas encore joué
    const [home, away] = dec(title).split(/\s+vs\s+/);
    const dm = (li.match(/(\d{2})\/(\d{2})/) || [])[0] || null;
    out.push({ live, home: (home || "").trim(), away: (away || "").trim(), dm });
  }
  return out;
}
const secs = gt => { const [m, s] = String(gt || "0:0").split(":").map(Number); return (m || 0) * 60 + (s || 0); };

export const handler = async (event) => {
  try {
    const q = event.queryStringParameters || {};
    if (q.team) {
      const key = norm(q.team);
      const games = parseGames(await getText(ROOT + "/saison-reguliere/calendrier"))
        .filter(g => norm(g.home).includes(key) || norm(g.away).includes(key) || key.includes(norm(g.home)) || key.includes(norm(g.away)));
      return { statusCode: 200, headers, body: JSON.stringify({ games }) };
    }
    if (q.live) {
      const id = String(q.live).replace(/\D/g, "");
      const r = await fetch(`${LIVE}/data/${id}/data.json`, { headers: UA });
      if (!r.ok) return { statusCode: 502, headers, body: JSON.stringify({ error: "LiveStats answered " + r.status }) };
      const d = await r.json();
      const teams = {};
      ["1", "2"].forEach(k => {
        const t = d.tm[k], starters = [], names = {};
        Object.values(t.pl || {}).forEach(p => { names[p.shirtNumber] = p.scoreboardName || p.name; if (p.starter) starters.push(String(p.shirtNumber)); });
        teams[k] = { name: t.name, score: Number(t.score) || 0, starters, names };
      });
      const ev = (d.pbp || []).map(x => {
        const rem = secs(x.gt), p = Number(x.period) || 1;
        const t = x.periodType === "OVERTIME" ? 2400 + (p - 1) * 300 + (300 - rem) : (p - 1) * 600 + (600 - rem);
        return [t, Number(x.actionNumber) || 0, Number(x.tno) || 0, String(x.shirtNumber || ""), x.actionType, x.subType || "", Number(x.s1) || 0, Number(x.s2) || 0, Number(x.previousAction) || 0];
      }).sort((a, b) => a[0] - b[0] || a[1] - b[1]);
      // Tirs à 2/3 pts avec position (tm[k].shot[]) : pour la carte de tir par zone.
      const shots = ["1", "2"].flatMap(k => ((d.tm[k] || {}).shot || [])
        .filter(s => s.x != null && s.y != null)
        .map(s => [Number(k), String(s.shirtNumber || ""), Number(s.r) || 0, Number(s.x), Number(s.y)]));
      return { statusCode: 200, headers: { ...headers, "Cache-Control": "public, max-age=600" }, body: JSON.stringify({ teams, ev, shots }) };
    }
    return { statusCode: 400, headers, body: JSON.stringify({ error: "Use ?team=<name> or ?live=<id>." }) };
  } catch (e) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: String(e.message || e) }) };
  }
};
