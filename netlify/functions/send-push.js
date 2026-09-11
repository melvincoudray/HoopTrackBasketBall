// Fonction serveur (Netlify Function) — appelée par l'app à chaque déclencheur (nouveau match
// importé, nouvelle ressource, message d'accueil modifié) pour envoyer une vraie notification
// push à tous les appareils abonnés de l'équipe concernée. L'envoi doit obligatoirement passer
// par un serveur (la clé privée VAPID ne doit jamais être exposée au navigateur).

import webpush from "web-push";
import { createClient } from "@supabase/supabase-js";

// Clé publique VAPID — identique à celle fixée en dur côté client (App.jsx), pour être certain
// qu'elles correspondent toujours exactement (un décalage entre les deux ferait échouer les
// envois silencieusement). Seule la clé PRIVÉE reste une variable d'environnement (secrète).
const VAPID_PUBLIC_KEY = "BBkeNtAnlqwbvvUBohClr7rqRoeFqo89NNDfOr24wDzI2Ebb-ssWKwyywfYj2nWGOXTkwhxe0-WpYMsbi47xw2Q";

export const handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method not allowed" };
  }

  const { teamId, title, body, url } = JSON.parse(event.body || "{}");
  if (!teamId || !title || !body) {
    return { statusCode: 400, body: "Missing teamId, title, or body" };
  }

  const VAPID_PRIVATE_KEY = process.env.VAPID_PRIVATE_KEY;
  if (!VAPID_PRIVATE_KEY) {
    return { statusCode: 500, body: "VAPID_PRIVATE_KEY not configured on the server" };
  }
  webpush.setVapidDetails("mailto:contact@hooptrack.app", VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY);

  const supabase = createClient(process.env.VITE_SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.VITE_SUPABASE_ANON_KEY);
  const { data: subs, error } = await supabase.from("push_subscriptions").select("endpoint, subscription").eq("team_id", teamId);
  if (error) {
    return { statusCode: 500, body: "Supabase error: " + error.message };
  }

  const payload = JSON.stringify({ title, body, url: url || "/" });
  let sent = 0, expired = 0;
  for (const row of subs || []) {
    try {
      await webpush.sendNotification(row.subscription, payload);
      sent++;
    } catch (err) {
      // Un abonnement expiré (410 Gone, ou 404) doit être retiré — l'appareil ne veut plus
      // recevoir de notifications (désinstallation, changement de navigateur, etc.).
      if (err.statusCode === 410 || err.statusCode === 404) {
        await supabase.from("push_subscriptions").delete().eq("endpoint", row.endpoint);
        expired++;
      }
    }
  }

  return { statusCode: 200, body: JSON.stringify({ sent, expired, total: (subs || []).length }) };
};
