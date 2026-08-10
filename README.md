# BahiSathi website

The real, deployable BahiSathi web app - marketing site, real login/signup
(Supabase Auth), and the bill-scanning ledger tool, all in one.

Unlike the earlier Claude.ai artifact version, this is a normal website and
can talk to Supabase directly without any sandbox restrictions.

## How bill scanning works here

The browser calls `POST /api/extract-bill` on your already-deployed WhatsApp
bot server (`tivaro-whatsapp-server` on Render), which safely holds the
Anthropic API key server-side. Make sure that server has been redeployed with
the updated `server.js` (the one with the `cors` package and the
`/api/extract-bill` route) before testing this site.

## Local development

```
npm install
npm run dev
```

## Deploying on Render

1. Push this folder to a new GitHub repo (e.g. `bahisathi-web`)
2. In Render: New + -> Static Site -> connect the `bahisathi-web` repo
3. Build command: `npm run build`
4. Publish directory: `dist`
5. Deploy

## Connecting your domain (bahisathi.in)

1. In Render, open your new static site -> Settings -> Custom Domains -> Add
   `bahisathi.in` (and `www.bahisathi.in` if you want both)
2. Render will show you a CNAME (or A record) to add
3. Go to wherever you registered bahisathi.in (GoDaddy, Namecheap, etc.) ->
   DNS settings -> add the record Render gave you
4. DNS changes can take anywhere from a few minutes to a few hours to
   propagate; Render will show "Verified" once it detects the domain pointing
   correctly, and will auto-provision an SSL certificate

## Known limitations for now

- Email confirmation is currently disabled on Supabase for faster testing -
  turn it back on (Authentication > Providers > Email > "Confirm email")
  before real users sign up, so people can't create accounts with emails
  they don't own
- The `/api/extract-bill` endpoint on the WhatsApp bot server currently has
  no rate limiting - fine for early testing, worth adding before wider
  public traffic to avoid abuse driving up your Anthropic API costs
