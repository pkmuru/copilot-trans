// index.js
import 'dotenv/config';
import { ConfidentialClientApplication } from '@azure/msal-node';

const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  RTA_MEETING_ID,
  POLL_MS = '2000',
  VERBOSE = '0'
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

// ---------- Helpers ----------
const sleep = (ms) => new Promise(res => setTimeout(res, ms));
const base = 'https://graph.microsoft.com/beta';

function transcriptsListUrl(meetingId) {
  return `${base}/copilot/communications/realtimeActivityFeed/meetings/${encodeURIComponent(meetingId)}/transcripts`;
}
function transcriptDetailUrl(meetingId, transcriptId) {
  return `${base}/copilot/communications/realtimeActivityFeed/meetings/${encodeURIComponent(meetingId)}/transcripts/${encodeURIComponent(transcriptId)}`;
}

async function fetchWithRetry(url, options = {}, maxAttempts = 5) {
  let attempt = 0;
  while (true) {
    attempt++;
    const res = await fetch(url, options);

    // Treat all non-2xx as errors (retry some)
    if (!res.ok) {
      // Retry for throttling/temporary
      if ((res.status === 429 || res.status === 503) && attempt < maxAttempts) {
        const ra = res.headers.get('retry-after');
        const wait = ra ? Number(ra) * 1000 : Math.min(32000, 1000 * 2 ** attempt);
        console.warn(`[${res.status}] throttled; retrying in ${wait}ms...`);
        await sleep(wait);
        continue;
      }

      const errText = await res.text().catch(() => '');
      throw new Error(`HTTP ${res.status} ${res.statusText}\n${errText}`);
    }

    return res;
  }
}

// Safely parse JSON if present. Returns { _empty: true } when body is empty.
async function safeJson(res) {
  const status = res.status;
  // No content (204) or Accepted (202) often have empty bodies
  if (status === 204 || status === 202) return { _empty: true, _status: status };

  const text = await res.text();          // don’t call res.json() directly
  if (!text || text.trim() === '') return { _empty: true, _status: status };

  const ct = res.headers.get('content-type') || '';
  if (!ct.includes('application/json')) {
    // Sometimes beta endpoints return oddly typed responses; expose raw text
    return { _raw: text, _status: status };
  }

  try {
    return JSON.parse(text);
  } catch (e) {
    // Help debug: show a snippet of what we received
    const snippet = text.length > 800 ? text.slice(0, 800) + '…' : text;
    throw new Error(`Failed to parse JSON (status ${status}). Body starts with:\n${snippet}`);
  }
}

async function listTranscripts(accessToken, meetingId) {
  const res = await fetchWithRetry(transcriptsListUrl(meetingId), {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json'
    }
  });
  if (VERBOSE === '1') {
    console.log(`[list] HTTP ${res.status} ${res.statusText}`);
  }
  return safeJson(res);
}

async function getTranscript(accessToken, meetingId, transcriptId) {
  const res = await fetchWithRetry(transcriptDetailUrl(meetingId, transcriptId), {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json'
    }
  });
  if (VERBOSE === '1') {
    console.log(`[detail:${transcriptId}] HTTP ${res.status} ${res.statusText}`);
  }
  return safeJson(res);
}

// Best-effort pretty printer (beta shapes can change)
function prettyPrintTranscript(obj) {
  const id = obj?.id ?? obj?.transcriptId ?? obj?.metadata?.id ?? 'unknown';
  const created = obj?.createdDateTime ?? obj?.startDateTime ?? obj?.timestamp ?? obj?.metadata?.createdDateTime;
  const lang = obj?.language ?? obj?.locale ?? obj?.metadata?.language;
  const speaker = obj?.speakerId ?? obj?.speaker?.id ?? obj?.participantId;
  const text =
    obj?.text ??
    obj?.content ??
    obj?.combinedText ??
    (Array.isArray(obj?.alternatives) ? obj.alternatives[0]?.text : undefined);

  const header = [
    `• id: ${id}`,
    created ? `created: ${created}` : null,
    lang ? `lang: ${lang}` : null,
    speaker ? `speaker: ${speaker}` : null,
  ].filter(Boolean).join(' | ');

  console.log(header || `• id: ${id}`);
  if (text) console.log(`  ${text}`);
}

// ---------- Poller ----------
async function run() {
  console.log('== Microsoft Graph RTA Transcript Poller ==');
  console.log(`Meeting: ${RTA_MEETING_ID}`);
  console.log(`Interval: ${POLL_MS} ms`);

  let token = await getToken();
  let tokenAcquiredAt = Date.now();
  const seen = new Set();

  while (true) {
    try {
      // refresh token roughly hourly
      if (Date.now() - tokenAcquiredAt > 50 * 60 * 1000) {
        token = await getToken();
        tokenAcquiredAt = Date.now();
      }

      const list = await listTranscripts(token, RTA_MEETING_ID);

      if (list?._empty) {
        if (VERBOSE === '1') console.log('[list] Empty body (202/204 or no data yet).');
        await sleep(Number(POLL_MS));
        continue;
      }
      if (list?._raw) {
        console.log('[list] Non-JSON body received (showing first 400 chars):');
        console.log(String(list._raw).slice(0, 400));
        await sleep(Number(POLL_MS));
        continue;
      }

      if (VERBOSE === '1') {
        console.log('~ list payload ~');
        console.dir(list, { depth: null });
      }

      const items = Array.isArray(list?.value) ? list.value : (Array.isArray(list) ? list : []);
      for (const t of items) {
        const id = t?.id ?? t?.transcriptId ?? t?.metadata?.id;
        if (!id || seen.has(id)) continue;
        seen.add(id);

        console.log('\nNEW transcript fragment:');
        prettyPrintTranscript(t);

        // Fetch detail (may also be empty)
        const detail = await getTranscript(token, RTA_MEETING_ID, id);
        if (detail?._empty) {
          if (VERBOSE === '1') console.log(`[detail:${id}] Empty body.`);
          continue;
        }
        if (detail?._raw) {
          console.log(`[detail:${id}] Non-JSON body (first 400 chars):`);
          console.log(String(detail._raw).slice(0, 400));
          continue;
        }
        if (VERBOSE === '1') {
          console.log('~ detail payload ~');
          console.dir(detail, { depth: null });
        }
        console.log('Detail snippet:');
        prettyPrintTranscript(detail);
      }
    } catch (err) {
      // 401? try to refresh once
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
