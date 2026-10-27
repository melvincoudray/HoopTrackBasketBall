// Fonction Netlify — feuilles de match (box scores) de nm1.ffbb.com pour une équipe.
//  • GET ?team=vitre                  -> { matches: [{ id, slug, date, home, away, scoreHome, scoreAway }] }
//                                         (matchs JOUÉS de l'équipe, du plus ancien au plus récent)
//  • GET ?team=vitre&match=<id>-<slug> -> { id, date, time, home, away, ours:{name,players,total}, opp:{...} }
//  • GET ?mode=league                  -> { matches: [{ id, slug, dm, home, away }] }  (TOUS les matchs joués de la ligue)
//  • GET ?mode=leaguematch&match=<id>-<slug> -> { id, date, teams: [{ name, total }] }  (totaux des 2 équipes, rebonds d'équipe inclus)
// Aucune dépendance externe (ESM, fetch natif de Node 18+).

const ROOT = "https://nm1.ffbb.com";
const UA = { "User-Agent": "HoopTrack/1.0 (box score sync)" };

function decodeEntities(s) {
  return String(s)
    .replace(/&nbsp;/g, " ").replace(/&eacute;/g, "é").replace(/&Eacute;/g, "É").replace(/&egrave;/g, "è").replace(/&amp;/g, "&").replace(/&quot;/g, '"')
    .replace(/&#0?39;|&apos;|&rsquo;|&#8217;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)));
}
const strip = s => decodeEntities(String(s).replace(/<[^>]*>/g, " ")).replace(/\s+/g, " ").trim();
export const norm = s => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]/g, "");

// Matchs joués listés dans le calendrier d'une page (même bloc "display-games" sur toutes les pages) :
// [{ id, slug, date: "JJ/MM", home, away, scoreHome, scoreAway }]
export function parseSchedule(html) {
  const out = new Map();
  const re = /<li class="display-games__third-list__entry[^"]*">([\s\S]*?)<\/li>/g;
  let m;
  while ((m = re.exec(html))) {
    const li = m[1];
    const link = li.match(/href="\/saison-reguliere\/match\/(\d+)-([^"]+)"[^>]*title="([^"]*)"/);
    if (!link) continue;
    const date = (strip((li.match(/<span class="date">([\s\S]*?)<\/span>/) || [])[1] || "").match(/(\d{2})\/(\d{2})/) || []);
    const alts = [...li.matchAll(/<img class="picture"[^>]*alt="([^"]*)"/g)].map(x => decodeEntities(x[1]).trim());
    const scores = [...li.matchAll(/<span class="team__score[^"]*">\s*([^<]*?)\s*<\/span>/g)].map(x => x[1].trim());
    const sh = /^\d+$/.test(scores[0] || "") ? Number(scores[0]) : null;
    const sa = /^\d+$/.test(scores[1] || "") ? Number(scores[1]) : null;
    const [home, away] = decodeEntities(link[3]).split(/\s+vs\s+/);
    if (!out.has(link[1])) out.set(link[1], { id: link[1], slug: link[2], dm: date[1] ? date[1] + "/" + date[2] : null, home: (alts[0] || home || "").trim(), away: (alts[1] || away || "").trim(), scoreHome: sh, scoreAway: sa });
  }
  return [...out.values()];
}


// Tous les matchs (joués ou à venir) du calendrier d'une page : + horaire et journée.
// [{ id, slug, dm: "JJ/MM", time: "HH:MM", journee: "Journée 8, Group B", home, away, played }]
export function parseCalendar(html) {
  const out = new Map();
  const re = /<li class="display-games__third-list__entry[^"]*">([\s\S]*?)<\/li>/g;
  let m;
  while ((m = re.exec(html))) {
    const li = m[1];
    const link = li.match(/href="\/saison-reguliere\/match\/(\d+)-([^"]+)"[^>]*title="([^"]*)"/);
    if (!link) continue;
    const dt = strip((li.match(/<span class="date">([\s\S]*?)<\/span>/) || [])[1] || "").match(/(\d{2})\/(\d{2})\s+(\d{2}:\d{2})/) || [];
    const jr = strip((li.match(/<span class="game__day">([\s\S]*?)<\/span>/) || [])[1] || "");
    const alts = [...li.matchAll(/<img class="picture"[^>]*alt="([^"]*)"/g)].map(x => decodeEntities(x[1]).trim());
    const scores = [...li.matchAll(/<span class="team__score[^"]*">\s*([^<]*?)\s*<\/span>/g)].map(x => x[1].trim());
    const [home, away] = decodeEntities(link[3]).split(/\s+vs\s+/);
    if (!out.has(link[1])) out.set(link[1], { id: link[1], slug: link[2], dm: dt[1] ? dt[1] + "/" + dt[2] : null, time: dt[3] || null, journee: jr, home: (alts[0] || home || "").trim(), away: (alts[1] || away || "").trim(), played: /^\d+$/.test(scores[0] || "") });
  }
  return [...out.values()];
}

const MOIS = ["janvier", "fevrier", "mars", "avril", "mai", "juin", "juillet", "aout", "septembre", "octobre", "novembre", "decembre"];
// "Derniers / Prochains matchs" d'une page d'équipe : salle des matchs à venir.
// [{ home, away, journee, venue }]
export function parseVenues(html) {
  const out = [];
  const re = /<li class="main__team__last-next-games__entry[^"]*">([\s\S]*?)<\/li>/g;
  let m;
  while ((m = re.exec(html))) {
    const li = m[1];
    const venue = strip((li.match(/fa-map-marker-alt"><\/i>([\s\S]*?)<\/div>/) || [])[1] || "");
    const teams = [...li.matchAll(/<a href="\/saison-reguliere\/equipe\/(\d+)-[^"]*" title="([^"]*)"/g)].map(x => decodeEntities(x[2]).trim());
    const jr = strip((li.match(/<span class="game__day[^"]*">([\s\S]*?)<\/span>\s*(?:<div|<\/span>)/) || [])[1] || "");
    const dd = strip((li.match(/<div class="date">([\s\S]*?)<\/div>/) || [])[1] || "").match(/(\d{1,2})\s+([a-zéèêûîôàç]+)\.?\s*-\s*(\d{2}:\d{2})/i) || [];
    const mi = MOIS.indexOf(norm(dd[2] || ""));
    const played = /last-next-games__entry__bottom--last/.test(li);
    if (teams.length >= 2) out.push({ home: teams[0], away: teams[1], journee: (jr.match(/Journée[^,]*,?[^<]*/) || [jr])[0].trim(), venue: played ? "" : venue, dm: mi >= 0 ? String(dd[1]).padStart(2, "0") + "/" + String(mi + 1).padStart(2, "0") : null, time: dd[3] || null, played });
  }
  return out;
}

// "Adresse : Salle de la Poultière, 3 bis rue de la Poultière 35500 Vitre" -> { hall, city }
export function parseAddress(html) {
  const a = strip((html.match(/<p class="team__infos">\s*Adresse\s*:([\s\S]*?)<\/p>/) || [])[1] || "");
  const city = (a.match(/\b\d{5}\s+(.+)$/) || [])[1] || "";
  return { hall: a.split(",")[0].trim(), city: city.trim(), address: a };
}

// "4-8" -> [4, 8, 4] (réussis, tentés, manqués) ; "" / "-" -> null
function made(cell) {
  const m = String(cell).match(/^(\d+)\s*-\s*(\d+)$/);
  return m ? { r: Number(m[1]), t: Number(m[2]) } : null;
}
const num = c => { const n = Number(String(c).replace(",", ".").trim()); return String(c).trim() !== "" && !isNaN(n) ? n : null; };

// Ligne du tableau -> stats brutes { min, pts, r2, t2, r3, t3, lfr, lft, ro, rd, rt, pd, in, bp, ct, fte, fpr, pm, eva }
function rowStats(cells) {
  // cellules attendues (avec N° et nom) : N°, Joueurs, Min, Pts, 2R-2T, 3R-3T, T%, LFR-LFT, LF%, RO, RD, RT, PD, IN, BP, CT, FTE, FPR, +/-, EVA
  const c = cells;
  const s2 = made(c[4]) || { r: 0, t: 0 }, s3 = made(c[5]) || { r: 0, t: 0 }, sf = made(c[7]) || { r: 0, t: 0 };
  return {
    min: num(c[2]) ?? 0, pts: num(c[3]) ?? 0, r2: s2.r, t2: s2.t, r3: s3.r, t3: s3.t, lfr: sf.r, lft: sf.t,
    ro: num(c[9]), rd: num(c[10]), rt: num(c[11]), pd: num(c[12]), in: num(c[13]), bp: num(c[14]), ct: num(c[15]),
    fte: num(c[16]), fpr: num(c[17]), pm: num(c[18]), eva: num(c[19]),
  };
}

// Page de match -> { date: "AAAA-MM-JJ", time, teams: [{ name, players:[{number,name,stats}], total, team }] }
export function parseMatch(html) {
  // L'en-tête du match est tout en bas de la page (après le calendrier de la ligue).
  const from = Math.max(0, html.indexOf("main__game"));
  const d = strip(html.slice(from, from + 15000)).match(/(\d{2})\/(\d{2})\/\s*(\d{4})\s+(\d{2}:\d{2})/);
  const date = d ? `${d[3]}-${d[2]}-${d[1]}` : null;
  const teams = [];
  const re = /<section class="main__game__table__entry">\s*<h2[^>]*>([\s\S]*?)<\/h2>([\s\S]*?)<\/section>/g;
  let m;
  while ((m = re.exec(html))) {
    const name = strip(m[1]);
    const body = m[2];
    const players = [];
    const tbody = (body.match(/<tbody>([\s\S]*?)<\/tbody>/) || [])[1] || "";
    let teamExtra = null;
    for (const tr of tbody.match(/<tr>[\s\S]*?<\/tr>/g) || []) {
      const cells = [...tr.matchAll(/<td[^>]*>([\s\S]*?)<\/td>/g)].map(x => strip(x[1]));
      if (/^[ÉE]quipe$/i.test(cells[0] || "")) { // rebonds d'équipe : colspan=2 => décalage de 1
        teamExtra = { ro: num(cells[8]) ?? 0, rd: num(cells[9]) ?? 0, rt: num(cells[10]) ?? 0 };
        continue;
      }
      if (cells.length < 20) continue;
      players.push({ number: cells[0], name: cells[1], stats: rowStats(cells) });
    }
    const foot = (body.match(/<tfoot>([\s\S]*?)<\/tfoot>/) || [])[1] || "";
    const fc = [...foot.matchAll(/<td[^>]*>([\s\S]*?)<\/td>/g)].map(x => strip(x[1]));
    // pied : "Totaux Équipe" (colspan=2) puis Min, Pts, ... => on réinsère une cellule pour retrouver les mêmes index
    const total = fc.length >= 19 ? rowStats([fc[0], "", ...fc.slice(1)]) : null;
    teams.push({ name, players, total, teamExtra });
  }
  return { date, time: d ? d[4] : null, teams };
}

async function getText(url) {
  const r = await fetch(url, { headers: UA });
  if (!r.ok) throw new Error("NM1 site answered " + r.status);
  return r.text();
}

export const handler = async (event) => {
  try {
    const q = event.queryStringParameters || {};
    const team = norm(q.team || "vitre");
    const season = (q.season || "2026").replace(/\D/g, "") || "2026";
    const headers = { "Content-Type": "application/json", "Cache-Control": "no-cache" };

    const lead0 = async () => getText(ROOT + "/saison-reguliere/leaders/equipes/" + season);
    const teamPathOf = (lead, name) => {
      for (const a of lead.match(/<a\b[^>]*>/gi) || []) {
        if (!/class="[^"]*\bteam\b[^"]*"/.test(a)) continue;
        const href = (a.match(/href="([^"]*\/equipe\/\d+-[^"]*)"/) || [])[1];
        const title = (a.match(/title="([^"]*)"/) || [])[1];
        if (href && title && norm(title) === norm(name)) return href;
      }
      return null;
    };

    // ?mode=games&team=vitre&opp=Rennes -> matchs (aller/retour) entre les deux équipes avec journée, date, heure, salle, ville
    if (q.mode === "games") {
      const lead = await lead0();
      const ownPath = teamPathOf(lead, team);
      if (!ownPath) return { statusCode: 404, headers, body: JSON.stringify({ error: "Team not found in the NM1 table." }) };
      const oppN = norm(q.opp || "");
      const ownHtml = await getText(ownPath.startsWith("http") ? ownPath : ROOT + ownPath);
      const oppPath = teamPathOf(lead, q.opp || "");
      const oppHtml = oppPath ? await getText(oppPath.startsWith("http") ? oppPath : ROOT + oppPath) : "";
      const pair = g => oppN && [g.home, g.away].some(n => norm(n) === oppN) && [g.home, g.away].some(n => norm(n) === team);
      const byKey = new Map();
      const put = (g) => { const k = norm(g.home) + "|" + norm(g.away); byKey.set(k, { ...(byKey.get(k) || {}), ...Object.fromEntries(Object.entries(g).filter(([, v]) => v !== null && v !== "" && v !== undefined)) }); };
      for (const html of [ownHtml, oppHtml]) {
        parseCalendar(html).filter(pair).forEach(put);
        parseVenues(html).filter(pair).forEach(v => put({ ...v, journee: v.journee.replace(/^NM1,\s*/, "") }));
      }
      const games = [...byKey.values()].sort((a, b) => String(a.dm || "").split("/").reverse().join("").localeCompare(String(b.dm || "").split("/").reverse().join("")));
      const cache = {};
      const addr = async (name) => {
        const k = norm(name);
        if (cache[k]) return cache[k];
        let html = ownHtml;
        if (k !== team) { const pth = teamPathOf(lead, name); html = pth ? await getText(pth.startsWith("http") ? pth : ROOT + pth) : ""; }
        return (cache[k] = parseAddress(html));
      };
      const res = [];
      for (const g of games) {
        const ad = await addr(g.home);
        res.push({ ...g, venue: g.venue || ad.hall || "", city: ad.city || "", hostsAddress: ad.address || "" });
      }
      return { statusCode: 200, headers, body: JSON.stringify({ games: res }) };
    }

    // ?mode=league -> tous les matchs joués de la ligue (page calendrier)
    if (q.mode === "league") {
      const html = await getText(ROOT + "/saison-reguliere/calendrier");
      const matches = parseCalendar(html).filter(g => g.played).map(g => ({ id: g.id, slug: g.slug, dm: g.dm, home: g.home, away: g.away }));
      return { statusCode: 200, headers, body: JSON.stringify({ fetchedAt: new Date().toISOString(), matches }) };
    }
    // ?mode=leaguematch&match=<id>-<slug> -> totaux des deux équipes (sans les joueurs : réponse légère)
    if (q.mode === "leaguematch" && q.match) {
      const mid = String(q.match).replace(/[^a-z0-9\-]/gi, "");
      const p = parseMatch(await getText(ROOT + "/saison-reguliere/match/" + mid));
      if (p.teams.length < 2 || !p.teams.every(t => t.total)) return { statusCode: 404, headers, body: JSON.stringify({ error: "No box score on this match page yet." }) };
      const teams = p.teams.map(t => {
        const tot = { ...t.total };
        // Rebonds d'équipe : le pied de tableau les contient en général ; sinon joueurs + ligne "Équipe".
        const sum = k => t.players.reduce((a, pl) => a + (pl.stats[k] || 0), 0) + (t.teamExtra ? (t.teamExtra[k] || 0) : 0);
        ["ro", "rd", "rt"].forEach(k => { tot[k] = Math.max(tot[k] || 0, sum(k)); });
        return { name: t.name, total: tot };
      });
      return { statusCode: 200, headers, body: JSON.stringify({ id: mid, date: p.date, teams }) };
    }

    if (q.match) {
      const mid = String(q.match).replace(/[^a-z0-9\-]/gi, "");
      const html = await getText(ROOT + "/saison-reguliere/match/" + mid);
      const p = parseMatch(html);
      if (p.teams.length < 2) return { statusCode: 404, headers, body: JSON.stringify({ error: "No box score on this match page yet." }) };
      const ours = p.teams.find(t => norm(t.name).includes(team) || team.includes(norm(t.name)));
      const opp = p.teams.find(t => t !== ours);
      if (!ours || !opp) return { statusCode: 404, headers, body: JSON.stringify({ error: "Team not found on this match page." }) };
      return { statusCode: 200, headers, body: JSON.stringify({ id: mid, date: p.date, time: p.time, ours, opp }) };
    }

    // Liste : page d'une équipe -> calendrier ; on garde les matchs joués de l'équipe demandée.
    const lead = await getText(ROOT + "/saison-reguliere/leaders/equipes/" + season);
    let teamPath = null;
    for (const a of lead.match(/<a\b[^>]*>/gi) || []) {
      if (!/class="[^"]*\bteam\b[^"]*"/.test(a)) continue;
      const href = (a.match(/href="([^"]*\/equipe\/\d+-[^"]*)"/) || [])[1];
      const title = (a.match(/title="([^"]*)"/) || [])[1];
      if (href && title && norm(title) === team) { teamPath = href; break; }
    }
    if (!teamPath) return { statusCode: 404, headers, body: JSON.stringify({ error: "Team not found in the NM1 table." }) };
    const html = await getText(teamPath.startsWith("http") ? teamPath : ROOT + teamPath);
    const matches = parseSchedule(html)
      .filter(m => m.scoreHome !== null && m.scoreAway !== null && (norm(m.home) === team || norm(m.away) === team))
      .map(m => ({ ...m, home: m.home, away: m.away }));
    return { statusCode: 200, headers, body: JSON.stringify({ fetchedAt: new Date().toISOString(), matches }) };
  } catch (e) {
    return { statusCode: 502, body: JSON.stringify({ error: String(e.message || e) }) };
  }
};
