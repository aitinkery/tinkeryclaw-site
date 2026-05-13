# Agent Anatomy

Five self-contained HTML experiences for teaching how AI agents actually work — using real data from Tinkery Bot, the AI running the AI Tinkery makerspace at Stanford GSE.

## Pages

| File | Experience | What it does |
|------|-----------|--------------|
| `index.html` | Landing | Links to all 5 experiences. Hero + card grid. |
| `turn.html` | Anatomy of a Turn | Step-through of one real $11.34 turn with tool call costs annotated. |
| `failure.html` | Watch It Fail | 3 real failure cases with dual-pane dialog + director's commentary. |
| `sandbox.html` | Try It | Mock agent sandbox — 8 pre-scripted example queries, simulated tool calls and costs. |
| `slider.html` | Autonomy Stages | Interactive slider through the 5-stage SAO framework from Andon Labs. |
| `ledger.html` | Live Ledger | Pedagogical version of the Tinkery Bot cost dashboard with explanation panels. |

## Running locally

```bash
cd ~/projects/agent-anatomy
python3 -m http.server 8000
# Open: http://localhost:8000
```

No build step. No dependencies. Vanilla HTML/CSS/JS.

## Data files

| File | Contents | Update frequency |
|------|----------|-----------------|
| `data/ledger-snapshot.json` | Full ledger data (cost, turns, commits by day/week/project) | Manual snapshot — re-copy from `~/.openclaw/workspace/ledger/ledger.json` |
| `data/outcomes-snapshot.json` | Bites analytics (sessions, activations, EXIF fills) | Manual snapshot — re-copy from `~/.openclaw/workspace/ledger/outcomes.json` |
| `data/turn-example.json` | The $11.34 EXIF feature turn, step-by-step | Hand-authored from real session data, 2026-05-08 |
| `data/failure-cases.json` | Three failure case studies with dialog + commentary | Hand-authored from real session logs |
| `data/stages.json` | Five SAO framework stages with descriptions | Stable — update only if the framework changes |

### Updating the ledger snapshot

```bash
cp ~/.openclaw/workspace/ledger/ledger.json ~/projects/agent-anatomy/data/ledger-snapshot.json
cp ~/.openclaw/workspace/ledger/outcomes.json ~/projects/agent-anatomy/data/outcomes-snapshot.json
```

Or trim commits per day to keep file size manageable:

```bash
cat ~/.openclaw/workspace/ledger/ledger.json | python3 -c "
import json, sys
d = json.load(sys.stdin)
for day in d['days']:
    for proj in day.get('commits_by_project', {}):
        day['commits_by_project'][proj] = day['commits_by_project'][proj][:3]
print(json.dumps(d, indent=2))
" > data/ledger-snapshot.json
```

## Design language

Matches the existing Tinkery Bot brand:

- **Background**: `#faf6f0` (warm cream)
- **Accent**: `#d97757` (terracotta orange)
- **Navy**: `#1a365d`
- **Headings**: Georgia serif
- **Body**: system sans-serif
- **Mono**: SF Mono / Menlo

All shared styles are in `style.css`.

## Adding a new failure case

Edit `data/failure-cases.json` and add a new object with:

```json
{
  "id": "unique-id",
  "title": "Short title",
  "date": "YYYY-MM-DD",
  "cost_wasted": 0.00,
  "summary": "One paragraph summary",
  "lesson": "One sentence lesson",
  "lesson_tag": "Where this lesson lives (MEMORY.md note, external link, etc.)",
  "dialog": [
    { "role": "user", "content": "..." },
    { "role": "agent", "content": "..." }
  ],
  "commentary": [
    { "at_step": 0, "type": "warning|error|critical|lesson|info", "text": "..." }
  ]
}
```

## Adding a mock scenario to the sandbox

Edit `sandbox.html` and add an entry to `MOCK_RESPONSES`:

```js
'your prompt keywords here': {
  steps: [
    { type: 'thinking', text: 'Agent reasoning...', cost: 0.03 },
    { type: 'tool', name: 'ToolName', args: 'arguments', result: 'tool output', cost: 0.04 },
    { type: 'answer', text: 'Final response text' }
  ],
  totalCost: 0.07,
  tokens: 800,
},
```

## Production sandbox (not deployed)

To deploy a real agent sandbox (instead of the mock), you'd need:

1. **Cloudflare Worker** at `/api/agent` — proxies requests to Anthropic API, holds the API key server-side
2. **Auth**: shared secret in Worker + localStorage on client
3. **Tools**: simulated (no real filesystem) — Write/Read/Search operate on a sandboxed in-memory object
4. **Cost cap**: Worker returns 402 if estimated cost > $0.10 and user hasn't confirmed
5. **Rate limit**: 1 request/minute tracked by IP in KV store

The mock approach is better for a teaching exhibit — it always works, can't be abused, and lets you hand-author exactly the tool call sequence you want to show.

## Placeholders and known gaps

- `turn-example.json`: costs are reconstructed from session logs, not pulled directly (session logs don't expose per-step cost). Values are accurate to within ~10%.
- `failure-cases.json` Case C (Bengt): reconstructed from Andon Labs' public blog post, not from a session log. Dialog is illustrative.
- `ledger-snapshot.json`: session logs before 2026-04-07 are permanently lost (OpenClaw reset event on 2026-04-21). All totals are floor estimates.

---

Built by Tinkery Bot · AI Tinkery at Stanford GSE · May 2026
