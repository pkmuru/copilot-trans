import 'dotenv/config';
import { ConfidentialClientApplication } from '@azure/msal-node';

const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  RTA_MEETING_ID,
  POLL_MS = '2000',           // adjust polling interval (ms)
  VERBOSE = '0'               // set to '1' to print raw JSON
} = process.env;

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !RTA_MEETING_ID) {
  console.error('Missing env. Set TENANT_ID, CLIENT_ID, CLIENT_SECRET, RTA_MEETING_ID in .env');
  process.exit(1);
}

// ---------- Auth (app-only) ----------
const cca = new ConfidentialClientApplication({
  auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}`, clientSecret: CLIENT_SECRET }
});

async function getToken() {
  const r = await cca.acquireTokenByClientCredential({ scopes: ['https://graph.microsoft.com/.default'] });
  if (!r?.accessToken) throw new Error('Failed to get app token');
  return r.accessToken;
}

// ---------- Fetch helpers with backoff ----------
const sleep = (ms) => new Promise(res => setTimeout(res, ms));

async function fetchWithRetry(url, options = {}, maxAttempts = 5) {
  let attempt = 0;
  while (true) {
    attempt++;
    const res = await fetch(url, options);
    if (res.ok) return res;

    // 429/503: honor Retry-After if present, else exponential backoff
    if ((res.status === 429 || res.status === 503) && attempt < maxAttempts) {
      const ra = res.headers.get('retry-after');
      const wait = ra ? Number(ra) * 1000 : Math.min(32000, 1000 * 2 ** attempt);
      console.warn(`[${res.status}] throttled; retrying in ${wait}ms...`);
      await sleep(wait);
      continue;
    }

    // Helpful error output
    const body = await res.text().catch(() => '');
    throw new Error(`HTTP ${res.status} ${res.statusText} for ${url}\n${body}`);
  }
}

// ---------- RTA endpoints ----------
const base = 'https://graph.microsoft.com/beta';
const transcriptsListUrl = (meetingId) =>
  `${base}/copilot/communications/realtimeActivityFeed/meetings/${encodeURIComponent(meetingId)}/transcripts`;
const transcriptDetailUrl = (meetingId, transcriptId) =>
  `${base}/copilot/communications/realtimeActivityFeed/meetings/${encodeURIComponent(meetingId)}/transcripts/${encodeURIComponent(transcriptId)}`;

// Try to extract something readable from unknown beta shapes
function prettyPrintTranscript(obj) {
  // Best effort guesswork:
  const id = obj.id ?? obj.transcriptId ?? obj?.metadata?.id ?? 'unknown';
  const created = obj.createdDateTime ?? obj.startDateTime ?? obj.timestamp ?? obj?.metadata?.createdDateTime;
  const lang = obj.language ?? obj.locale ?? obj?.metadata?.language;
  const speaker = obj.speakerId ?? obj?.speaker?.id ?? obj.participantId;
  const text = obj.text ?? obj?.content ?? obj?.combinedText ?? obj?.alternatives?.[0]?.text;

  const header = [
    `• id: ${id}`,
    created ? `created: ${created}` : null,
    lang ? `lang: ${lang}` : null,
    speaker ? `speaker: ${speaker}` : null,
  ].filter(Boolean).join(' | ');

  console.log(header || `• id: ${id}`);
  if (text) console.log(`  ${text}`);
}

async function listTranscripts(accessToken, meetingId) {
  const res = await fetchWithRetry(transcriptsListUrl(meetingId), {
    headers: { Authorization: `Bearer ${accessToken}`, 'Accept': 'application/json' }
  });
  return res.json();
}

async function getTranscript(accessToken, meetingId, transcriptId) {
  const res = await fetchWithRetry(transcriptDetailUrl(meetingId, transcriptId), {
    headers: { Authorization: `Bearer ${accessToken}`, 'Accept': 'application/json' }
  });
  return res.json();
}

// ---------- Poller ----------
async function run() {
  console.log('== Microsoft Graph RTA Transcript Poller ==');
  console.log(`Meeting: ${RTA_MEETING_ID}`);
  console.log(`Interval: ${POLL_MS} ms`);
  const seen = new Set();
  let token = await getToken();
  let tokenAcquiredAt = Date.now();

  while (true) {
    try {
      // refresh token every 50 minutes (belt & braces)
      if (Date.now() - tokenAcquiredAt > 50 * 60 * 1000) {
        token = await getToken();
        tokenAcquiredAt = Date.now();
      }

      const list = await listTranscripts(token, RTA_MEETING_ID);
      const items = Array.isArray(list?.value) ? list.value : (Array.isArray(list) ? list : []);
      if (VERBOSE === '1') {
        console.log('~ list payload ~');
        console.dir(list, { depth: null });
      }

      // Process new items (dedupe by id)
      for (const t of items) {
        const id = t.id ?? t.transcriptId ?? t?.metadata?.id;
        if (!id || seen.has(id)) continue;
        seen.add(id);

        // Print summary row from list object
        console.log('\nNEW transcript fragment detected:');
        prettyPrintTranscript(t);

        // Fetch full detail
        try {
          const detail = await getTranscript(token, RTA_MEETING_ID, id);
          if (VERBOSE === '1') {
            console.log('~ detail payload ~');
            console.dir(detail, { depth: null });
          }
          // Try to print something human-friendly from detail too
          if (detail) {
            console.log('Detail snippet:');
            prettyPrintTranscript(detail);
          }
        } catch (e) {
          console.warn(`Failed to fetch detail for ${id}: ${e.message}`);
        }
      }
    } catch (err) {
      // Token expiry or permission issues – try a quick refresh once
      if (String(err.message || '').includes('401')) {
        console.warn('401 from Graph. Refreshing token...');
        token = await getToken();
        tokenAcquiredAt = Date.now();
      } else {
        console.warn('Poll error:', err.message);
      }
    }

    await sleep(Number(POLL_MS));
  }
}

run().catch(e => {
  console.error(e);
  process.exit(1);
});
