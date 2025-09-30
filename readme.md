# Node 18+ recommended (uses built-in fetch)
mkdir rta-transcripts && cd rta-transcripts
npm init -y
npm i @azure/msal-node dotenv

# create files
printf "%s\n" \
'TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_SECRET=your_client_secret
RTA_MEETING_ID=your-realtimeActivityMeeting-id' > .env
