# CELLIX Technical Guide (legacy)

> **Superseded.** This file is a historical GST/compliance-oriented draft.
>
> **Source of truth for Cellix-2026:** [`docs/CELLIX_TECHNICAL_DOCUMENTATION.md`](../../docs/CELLIX_TECHNICAL_DOCUMENTATION.md) (updated **July 18, 2026**).
>
> **Architecture diagrams:** [`docs/cellix-architecture-diagrams.html`](../../docs/cellix-architecture-diagrams.html)

## Current project map (quick pointer)

| Area | Where it lives now |
|------|--------------------|
| Excel add-in | `frontend/` — React 18 + Vite + Office.js, Ask/Plan/Action modes, preview Accept/Reject |
| NestJS API | `cellix_backend/` — Fastify, SSE conversation, complexity tiers 0–3, agents |
| Ops Dashboard | `Dashboard/` — Next.js :3100, request/planner log viewers |
| Specs | `specs/00`–`10` (tiering, modes, domain tools, critical bugfixes) |

### Write path (today)

```
POST /excel-ai/conversation
  → LlmRouter (shortcut | data | export | write | ask)
  → route=write → ENABLE_COMPLEXITY_TIERING
      → Tier 0 deterministic | Tier 1 single LLM | Tier 2 Gen→Verify | Tier 3 Planner→Executor→Verifier
  → normalizeExecutorOutput (A1 range → row/col for FORMAT-class)
  → sanitizeAction (indices required for FORMAT_RANGE)
  → ChangeSet preview → SSE actions
  → Add-in previewManager → Accept → Office.js
```

### Do not use this legacy draft for

- Redis / BullMQ / JWT auth (not implemented)
- Rules-engine GST workflows as the primary product path
- pnpm-only or Anthropic-only setup (project uses npm + OpenRouter)

See the main technical doc for env vars, module graph, SSE events, and testing.
