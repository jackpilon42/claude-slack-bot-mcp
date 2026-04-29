# Claude Slack Bot Startup

Use this after rebooting your Mac:

```bash
cd /Users/jackpilon/claude-slack-bot
npm run startup
```

What it does:

1. Loads `.env`
2. Validates required keys
3. Reuses an existing ngrok tunnel (or starts one)
4. Prints the Slack Request URL (`.../slack/events`)
5. Starts the bot

If the URL changed, update it in Slack:

1. Go to `https://api.slack.com/apps`
2. Open your app
3. Click `Event Subscriptions`
4. Paste the printed Request URL
5. Save

Keep the terminal open while the bot is running.

## Status Dashboard

To quickly see where startup is failing (env, bot process, ngrok, Slack reachability):

```bash
cd /Users/jackpilon/claude-slack-bot
npm run status
```

It prints:

- `.env` checks
- bot local port check (`localhost:3000`)
- `/slack/events` local check
- ngrok tunnel URL
- recommended Slack Request URL
- outbound `slack.com` reachability
