import 'dotenv/config';
import { ClientSecretCredential } from '@azure/identity';
import { AzureIdentityAuthenticationProvider } from '@microsoft/kiota-authentication-azure';
import { FetchRequestAdapter } from '@microsoft/kiota-http-fetchlibrary';
import { createBaseAgentsM365CopilotBetaServiceClient } from '@microsoft/agents-m365copilot-beta';
import { RequestInformation, HttpMethod } from '@microsoft/kiota-abstractions';

const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  RTA_MEETING_ID,
  POLL_MS = '2000',
  VERBOSE = '0'
} = process.env;

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !RTA_MEETING_ID) {
  console.error('Set TENANT_ID, CLIENT_ID, CLIENT_SECRET, RTA_MEETING_ID in .env');
  process.exit(1);
}

/**
 * Auth & client setup (app-only):
 * - Kiota requires an authentication provider and a request adapter.
 * - Use /.default scope for application permissions with Graph.
 */
const credential = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
const authProvider = new AzureIdentityAuthenticationProvider(
  credential,
  ['https://graph.microsoft.com/.default'] // app-only token
);
const adapter = new FetchRequestAdapter(authProvider);
adapter.baseUrl = 'https://graph.microsoft.com/beta';

// Creating the typed client (in case you want to explore typed builders via intellisense)
const client = createBaseAgentsM365CopilotBetaServiceClient(adapter);

/**
 * Some SDKs may not yet expose typed request builders for brand-new beta routes.
 * To be future-proof, we use Kiota's low-level RequestInformation to call:
 *   GET /copilot/communications/realtimeActivityFeed/meetings/{realtimeActivityMeeting-id}/transcripts
 *   GET /copilot/communications/realtimeActivityFeed/meetings/{id}/transcripts/{realTimeTranscript-id}
 */

async function listTranscriptsRaw(meetingId) {
  const ri = new RequestInformation();
  ri.urlTemplate = '{+baseurl}/copilot/communications/realtimeActivityFeed/meetings/{realtimeActivityMeeting-id}/transcripts';
  ri.pathParameters = { 'realtimeActivityMeeting-id': meetingId };
  ri.httpMethod = HttpMethod.GET;
  ri.headers['Accept'] = 'application/json';
  // returns any; shape can evolve in beta
  return adapter.sendPrimitiveAsync(ri, 'application/json');
}

async function getTranscriptRaw(meetingId, transcriptId) {
  const ri = new RequestInformation();
  ri.urlTemplate = '{+baseurl}/copilot/communications/realtimeActivityFeed/meetings/{realtimeActivityMeeting-id}/transcripts/{realTimeTranscript-id}';
  ri.pathParameters = { 'realtimeActivityMeeting-id': meetingId, 'realTimeTranscript-id': transcriptId };
  ri.httpMethod = HttpMethod.GET;
  ri.headers['Accept'] = 'application/json';
  return adapter.sendPrimitiveAsync(ri, 'application/json');
}

// best-effort pretty printer (beta payloads can change)
function pretty(obj) {
  const id = obj?.id ?? obj?.transcriptId ?? obj?.metadata?.id ?? 'unknown';
  const ts = obj?.createdDateTime ?? obj?.startDateTime ?? obj?.timestamp ?? obj?.metadata?.createdDateTime;
  const lang = obj?.language ?? obj?.locale ?? obj?.metadata?.language;
  const speaker = obj?.speakerId ?? obj?.speaker?.id ?? obj?.participantId;
  const text =
    obj?.text ??
    obj?.content ??
    obj?.combinedText ??
    (Array.isArray(obj?.alternatives) ? obj.alternatives[0]?.text : undefined);
  console.log(`• id: ${id}${ts ? ' | created: ' + ts : ''}${lang ? ' | lang: ' + lang : ''}${speaker ? ' | speaker: ' + speaker : ''}`);
  if (text) console.log('  ' + text);
}

async function run() {
  console.log('== RTA transcript poller (Copilot TS client + Kiota) ==');
  console.log(`Meeting: ${RTA_MEETING_ID}`);
  console.log(`Interval: ${POLL_MS} ms`);

  const seen = new Set();

  while (true) {
    try {
      const list = await listTranscriptsRaw(RTA_MEETING_ID);
      if (VERBOSE === '1') {
        console.log('~ list payload ~');
        console.dir(list, { depth: null });
      }

      const items = Array.isArray(list?.value) ? list.value : (Array.isArray(list) ? list : []);
      for (const t of items) {
        const tid = t?.id ?? t?.transcriptId ?? t?.metadata?.id;
        if (!tid || seen.has(tid)) continue;
        seen.add(tid);

        console.log('\nNEW transcript fragment:');
        pretty(t);

        try {
          const detail = await getTranscriptRaw(RTA_MEETING_ID, tid);
          if (VERBOSE === '1') {
            console.log('~ detail payload ~');
            console.dir(detail, { depth: null });
          }
          console.log('Detail snippet:');
          pretty(detail);
        } catch (e) {
          console.warn(`Failed to fetch detail for ${tid}: ${e.message}`);
        }
      }
    } catch (e) {
      console.warn('Poll error:', e.message);
    }
    await new Promise(r => setTimeout(r, Number(POLL_MS)));
  }
}

run();
