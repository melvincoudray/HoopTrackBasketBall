// Fonction planifiée (Netlify Scheduled Function) — tourne toutes les 15 minutes, vérifie les
// événements du planning de TOUTES les équipes, et envoie :
//   1. une notification 2h avant le début de chaque événement,
//   2. une notification au moment exact où l'événement commence.
// Chaque type de rappel n'est envoyé qu'une seule fois par événement (table
// "sent_event_reminders"), même si cette fonction tourne plusieurs fois avant l'heure visée.

import webpush from "web-push";
import { createClient } from "@supabase/supabase-js";

// Fenêtre de tolérance : la fonction tourne toutes les 15 min, donc un événement "tombe" dans
// la fenêtre d'un rappel à un moment ou un autre dans les 15 minutes qui suivent le passage au
// moment théorique du rappel — élargi ici à 20 minutes pour absorber tout retard d'exécution.
const WINDOW_MINUTES = 20;
// Identique à celle fixée en dur côté client (App.jsx) et dans send-push.js.
const VAPID_PUBLIC_KEY = "BBkeNtAnlqwbvvUBohClr7rqRoeFqo89NNDfOr24wDzI2Ebb-ssWKwyywfYj2nWGOXTkwhxe0-WpYMsbi47xw2Q";

// BUG RÉEL CORRIGÉ (signalé par l'utilisateur : un événement à 14h — heure française — déclenchait
// le rappel "2h avant" à 14h au lieu de 12h) : le serveur Netlify tourne en UTC, alors que les
// heures saisies dans le planning sont toujours en heure de Paris (été ou hiver). Sans cette
// conversion explicite, "14:00" était compris comme 14h UTC (= 16h à Paris), décalant tout de 2h.
// Calcule le décalage réel Paris/UTC pour une date donnée (gère automatiquement l'heure d'été
// UTC+2 et l'heure d'hiver UTC+1, sans dépendre d'une bibliothèque externe).
function getParisOffsetMinutes(date) {
  const utcDate = new Date(date.toLocaleString("en-US", { timeZone: "UTC" }));
  const parisDate = new Date(date.toLocaleString("en-US", { timeZone: "Europe/Paris" }));
  return (parisDate - utcDate) / 60000;
}
function parisTimeToUTC(dateStr, timeStr) {
  const naive = new Date(`${dateStr}T${timeStr}:00Z`); // traité provisoirement comme si "Z" (UTC)
  const offsetMinutes = getParisOffsetMinutes(naive);
  return new Date(naive.getTime() - offsetMinutes * 60000);
}

// Les deux rappels possibles pour un même événement : combien de minutes avant son début, le
// suffixe à ajouter à sa clé de déduplication, et comment construire le message envoyé.
const REMINDER_KINDS = [
  {
    minutesBefore: 120,
    suffix: ":2h",
    buildPayload: (ev) => ({
      title: `In 2h: ${ev.title}`,
      body: `${ev.startTime} · ${ev.type}${ev.location ? " · " + ev.location : ""}`,
      url: "/",
    }),
  },
  {
    minutesBefore: 0,
    suffix: ":start",
    buildPayload: (ev) => ({
      title: `Starting now: ${ev.title}`,
      body: `${ev.type}${ev.location ? " · " + ev.location : ""}`,
      url: "/",
    }),
  },
];

export const handler = async () => {
  const VAPID_PRIVATE_KEY = process.env.VAPID_PRIVATE_KEY;
  if (!VAPID_PRIVATE_KEY) {
    return { statusCode: 500, body: "VAPID_PRIVATE_KEY not configured on the server" };
  }
  webpush.setVapidDetails("mailto:contact@hooptrack.app", VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY);
  const supabase = createClient(process.env.VITE_SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.VITE_SUPABASE_ANON_KEY);

  // Récupère les événements de planning de CHAQUE équipe — stockés sous la clé
  // "team_<id>:planning_events" dans app_storage.
  const { data: rows, error } = await supabase.from("app_storage").select("key, value").like("key", "team_%:planning_events");
  if (error) return { statusCode: 500, body: "Supabase error: " + error.message };

  const now = new Date();
  console.log("[reminders] run started at", now.toISOString(), "— found", (rows || []).length, "team(s) with planning events");
  let notified = 0;

  for (const row of rows || []) {
    const teamId = row.key.replace(/^team_/, "").replace(/:planning_events$/, "");
    const events = Array.isArray(row.value) ? row.value : [];
    console.log("[reminders] team", teamId, "—", events.length, "event(s) in planning");
    for (const ev of events) {
      if (!ev.date || !ev.startTime || !ev.id) {
        console.log("[reminders]   event skipped (missing date/startTime/id):", JSON.stringify(ev));
        continue;
      }
      const startsAt = parisTimeToUTC(ev.date, ev.startTime);
      const minutesUntilStart = (startsAt - now) / 60000;
      console.log(`[reminders]   event "${ev.title}" (${ev.id}) — ${ev.date} ${ev.startTime} Paris → ${startsAt.toISOString()} UTC — ${minutesUntilStart.toFixed(1)} min until start`);

      for (const kind of REMINDER_KINDS) {
        const minutesUntilReminder = minutesUntilStart - kind.minutesBefore;
        // L'événement "entre" dans la fenêtre de ce rappel si on est entre 0 et WINDOW_MINUTES
        // minutes APRÈS le moment théorique du rappel (jamais avant, jamais trop après).
        if (minutesUntilReminder > 0 || minutesUntilReminder < -WINDOW_MINUTES) {
          console.log(`[reminders]     ${kind.suffix} — hors fenêtre (${minutesUntilReminder.toFixed(1)} min par rapport au moment visé) → ignoré`);
          continue;
        }

        const eventId = "reminder:" + ev.id + kind.suffix;
        const { data: already } = await supabase.from("sent_event_reminders").select("event_id").eq("event_id", eventId).maybeSingle();
        if (already) {
          console.log(`[reminders]     ${kind.suffix} — déjà envoyé précédemment → ignoré`);
          continue;
        }

        const { data: subs } = await supabase.from("push_subscriptions").select("endpoint, subscription").eq("team_id", teamId);
        console.log(`[reminders]     ${kind.suffix} — DANS la fenêtre, envoi à ${(subs || []).length} abonnement(s)…`);
        const payload = JSON.stringify(kind.buildPayload(ev));
        for (const sub of subs || []) {
          try {
            await webpush.sendNotification(sub.subscription, payload);
            console.log(`[reminders]       envoyé avec succès à ${sub.endpoint.slice(-30)}`);
          } catch (err) {
            console.error(`[reminders]       ÉCHEC d'envoi à ${sub.endpoint.slice(-30)} —`, err.statusCode, err.message);
            if (err.statusCode === 410 || err.statusCode === 404) {
              await supabase.from("push_subscriptions").delete().eq("endpoint", sub.endpoint);
            }
          }
        }
        await supabase.from("sent_event_reminders").insert({ event_id: eventId });
        notified++;
      }
    }
  }

  console.log("[reminders] run finished — total notified:", notified);
  return { statusCode: 200, body: JSON.stringify({ notified }) };
};
