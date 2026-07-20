# Domain tools — code review checklist (Spec 06)

Use this on every PR that touches `src/domain-tools/`.

- [ ] Tool is a plain TypeScript function (`DomainTool<TIn, TOut>`) — no Nest injectable with LLM deps
- [ ] No import of `OpenRouterService`, OpenAI SDK, Anthropic SDK, or any chat/completion client
- [ ] `DomainToolResult` always returns `confidence`, `exceptions`, and `sourceRefs` (non-optional)
- [ ] Fuzzy / low-confidence matches use `severity: 'flag'` (or `'block'`) — never silent auto-accept
- [ ] Executor must write **formulas** referencing tool outputs — never hard-coded numeric literals from `result.data`
- [ ] Ingestion fixtures are synthetic/redacted only — no real client documents
- [ ] Stub tools must not be exposed to end users as compliance-correct results
