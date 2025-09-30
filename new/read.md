

# Node 18+ (uses fetch). New folder:
mkdir rta-transcripts-sdk && cd rta-transcripts-sdk
npm init -y

# deps: Copilot client + Kiota auth/HTTP + Azure Identity + env
npm i @microsoft/agents-m365copilot-beta \
      @microsoft/kiota-authentication-azure \
      @microsoft/kiota-http-fetchlibrary \
      @azure/identity dotenv


TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_SECRET=your_client_secret
RTA_MEETING_ID=your-realtimeActivityMeeting-id
POLL_MS=2000
VERBOSE=0
