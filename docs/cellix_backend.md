# CELLIX Backend Notes (legacy)

> **Superseded.** Sections below describe a planned GST/rules-engine stack (Redis, BullMQ, JWT) that is **not** what Cellix-2026 ships.
>
> **Source of truth:** [`docs/CELLIX_TECHNICAL_DOCUMENTATION.md`](../../docs/CELLIX_TECHNICAL_DOCUMENTATION.md) (July 18, 2026).  
> **Diagrams:** [`docs/cellix-architecture-diagrams.html`](../../docs/cellix-architecture-diagrams.html).  
> **Rollout flag:** [`docs/COMPLEXITY_TIERING_ROLLOUT.md`](../../docs/COMPLEXITY_TIERING_ROLLOUT.md).

## What is actually running

| Topic | Current Cellix-2026 |
|-------|---------------------|
| Runtime | NestJS 11 + **Fastify**, port **4001** |
| Package manager | **npm** (not pnpm-required) |
| LLM | **OpenRouter** (`OPENROUTER_API_KEY`), tiered LOW/MEDIUM/HIGH models |
| DB | MongoDB via Mongoose — conversations, change_sets, audit, `request_logs`, `planner_logs` |
| Auth | **Not implemented** — open in local/dev |
| Cache / queues | No Redis/BullMQ in the shipped tree |
| Main API | `POST /excel-ai/conversation` (SSE) |
| Write dispatch | Complexity tiers 0–3 behind `ENABLE_COMPLEXITY_TIERING` |
| Agents | Planner / Executor / Verifier under `src/agents/` (Tier 3) |
| Logging | `logs/requests.log` + `logs/planner.log` (24h) + Mongo TTL 3d; Dashboard on :3100 |

## Bootstrap (accurate)

```bash
cd cellix_backend
npm install
# .env: PORT, MONGODB_URL, OPENROUTER_API_KEY, ENABLE_COMPLEXITY_TIERING
npm run start:dev
```

Module wiring lives in `src/app.module.ts` (config, database, logging, audit, health, excel-ai, sheets, domain-tools). Prefer the main technical doc §8 for the full module graph and §9 for API reference.

## Historical draft

The remainder of older versions of this file covered AnalyseModule / ClassifierModule / RulesEngine / Redis / JWT scaffolding. Treat those as **design history**, not the current codebase.
