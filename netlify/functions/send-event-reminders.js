// Fonction planifiée (Netlify Scheduled Function) — tourne toutes les 15 minutes, vérifie les
// événements du planning de TOUTES les équipes, et envoie une notification 2h avant le début de
// chacun. Chaque événement n'est notifié qu'une seule fois (table "sent_event_reminders"),
// même si cette fonction tourne plusieurs fois avant l'heure du rappel.

import webpush from "web-push";
import { createClient } from "@supabase/supabase-js";

const REMINDER_HOURS_BEFORE = 2;
// Fenêtre de tolérance : la fonction tourne toutes les 15 min, donc un événement "tombe" dans
// la fenêtre de rappel à un moment ou un autre dans les 15 minutes qui suivent son passage à
// "il reste exactement 2h" — élargi ici à 20 minutes pour absorber tout retard d'exécution.
const WINDOW_MINUTES = 20;
// Identique à celle fixée en dur côté client (App.jsx) et dans send-push.js.
const VAPID_PUBLIC_KEY = "BBkeNtAnlqwbvvUBohClr7rqRoeFqo89NNDfOr24wDzI2Ebb-ssWKwyywfYj2nWGOXTkwhxe0-WpYMsbi47xw2Q";

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
  let notified = 0;

  for (const row of rows || []) {
    const teamId = row.key.replace(/^team_/, "").replace(/:planning_events$/, "");
    const events = Array.isArray(row.value) ? row.value : [];
    for (const ev of events) {
      if (!ev.date || !ev.startTime || !ev.id) continue;
      const startsAt = new Date(`${ev.date}T${ev.startTime}:00`);
      const minutesUntilStart = (startsAt - now) / 60000;
      const minutesUntilReminder = minutesUntilStart - REMINDER_HOURS_BEFORE * 60;
      // L'événement "entre" dans la fenêtre de rappel si on est entre 0 et WINDOW_MINUTES
      // minutes APRÈS le moment théorique du rappel (jamais avant, jamais trop après).
      if (minutesUntilReminder > 0 || minutesUntilReminder < -WINDOW_MINUTES) continue;

      const eventId = "reminder:" + ev.id;
      const { data: already } = await supabase.from("sent_event_reminders").select("event_id").eq("event_id", eventId).maybeSingle();
      if (already) continue;

      const { data: subs } = await supabase.from("push_subscriptions").select("endpoint, subscription").eq("team_id", teamId);
      const payload = JSON.stringify({
        title: `In 2h: ${ev.title}`,
        body: `${ev.startTime} · ${ev.type}${ev.location ? " · " + ev.location : ""}`,
        url: "/",
      });
      for (const sub of subs || []) {
        try {
          await webpush.sendNotification(sub.subscription, payload);
        } catch (err) {
          if (err.statusCode === 410 || err.statusCode === 404) {
            await supabase.from("push_subscriptions").delete().eq("endpoint", sub.endpoint);
          }
        }
      }
      await supabase.from("sent_event_reminders").insert({ event_id: eventId });
      notified++;
    }
  }

  return { statusCode: 200, body: JSON.stringify({ notified }) };
};
