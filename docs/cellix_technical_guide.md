### 1. System Overview

CELLIX is an AI-powered Excel assistant implemented as a Microsoft Excel add-in using Office.js task pane architecture. It integrates a React 18 frontend with a NestJS backend that orchestrates an LLM multi-agent pipeline, deterministic routing (find/export, shortcuts), and a MongoDB database for conversations, audit trails, and change sets.

> **Note:** This document is a legacy GST/compliance-oriented draft. For the current Cellix-2026 architecture, see [`docs/CELLIX_TECHNICAL_DOCUMENTATION.md`](../../docs/CELLIX_TECHNICAL_DOCUMENTATION.md) (updated June 2026: LLM Router, tiered TOON, data aggregation vs find, sheet mentions, Tally header detection).

**Key components:**

- React frontend (Office.js task pane iframe)
- Node.js backend managing LLM, GST rules, and database
- MongoDB for audit and session data
- Named process mapping for system clarity

### 2. Excel Add-in Architecture

CELLIX is implemented as an Office.js task pane add-in embedded within Excel. The task pane operates inside a browser iframe and hosts the React application. Office.js APIs provide a bridge between the React UI and the Excel workbook, enabling reading and writing of cell ranges, and responding to selection changes.

#### 2.1 Why Excel, not Google Sheets

CELLIX targets the Indian Chartered Accountant (CA) market, where enterprise and mid-sized firms predominantly use Excel for compliance workflows. Key reasons include:

- CA firms’ audit files are typically in `.xlsx` format.
- Tally accounting software exports financials to Excel.
- The GST portal supports Excel imports.
- Google Sheets is mostly used by smaller solo practices; the industry standard is Excel.
- CELLIX V1 is therefore Excel-exclusive.

#### 2.2 Office.js task pane setup

The task pane setup leverages Office.js to enable seamless interaction with workbook data and user interface rendering within Excel.

#### 2.3 Frontend package choices

The frontend stack centers on React 18, ensuring a modern, responsive UI suited for the task pane environment.

### 3. Process: PROMPT_CAPTURE

This is the entry point process triggered when the CA clicks Run (or presses ⌘↵) in the sidebar. PROMPT_CAPTURE performs two essential functions:

- Captures the user’s prompt text from the textarea.
- Calls RANGE_READER to collect the current workbook context including the active selection.

### 4. Process: RANGE_READER

RANGE_READER uses Office.js Excel APIs to extract data from the user’s current selection. It retrieves:

- Cell values
- Cell formats
- Range address

Because Office.js runs client-side (in the browser context), this read operation is fast and responsive, unlike Google Apps Script which is server-side.

### 5. Process: PAYLOAD_COMPRESSOR

PAYLOAD_COMPRESSOR operates on the frontend before sending data to the backend. Its role is to minimize payload size and optimize token usage for the LLM by normalizing and compressing Excel data, which is often verbose.

**Main functions:**

- Strip formatting: e.g., convert "₹14,200.00" strings to numeric $14200$.
- Remove empty rows to avoid sending unnecessary data.
- Truncate data to a maximum of 50 rows due to LLM context window limits.
- Use compact JSON serialization (array-of-arrays rather than array-of-objects) to save approximately 30% tokens.
- Replace $null$ or $undefined$ cells with empty strings to preserve array shape.

### 6. Process: INTENT_CLASSIFIER

After authentication and credit checks, INTENT_CLASSIFIER runs on the backend. It analyzes the user prompt and column headers from PAYLOAD_COMPRESSOR to classify the request into one of five predefined workflows. This classification determines:

- Which system prompt template to load
- Which deterministic rules engine module to invoke
- Which LLM model tier MODEL_SELECTOR will select

### 7. Process: RULES_ENGINE

RULES_ENGINE applies deterministic checks on every row using rules sourced from MongoDB. It handles all decisions that do not require LLM inference, such as:

- Validating known HSN rates
- Checking blocked ITC categories
- Applying GSTR-2A matching tolerances

Rows passing through RULES_ENGINE cleanly do not consume LLM token budget, significantly reducing operational costs.

### 8. Process: MODEL_SELECTOR

MODEL_SELECTOR determines which LLM tier to use for rows that remain ambiguous after RULES_ENGINE processing. It bases its decision on:

- Workflow complexity
- Number of ambiguous rows
- Presence of high-severity edge cases flagged by RULES_ENGINE

#### 8.1 Model tiers and when to use each

CELLIX uses multiple model tiers with differing cost and capability profiles. Simpler cases use cheaper models, while complex or high-risk cases trigger more powerful and costly models, optimizing cost-efficiency.

### 9. Process: PROMPT_ASSEMBLER

PROMPT_ASSEMBLER builds the system prompt sent to the LLM from four layered components:

- **Layer 1 (Role + output contract, ~200 tokens):** Defines model identity, enforces JSON-only output, and sets patch schema. Sent on every call.
- **Layer 2 (GST rule tables, ~2000 tokens, cached):** Includes HSN rate tables, SAC tables, blocked ITC lists, and GSTR-2A tolerances loaded from MongoDB. Most valuable for caching.
- **Layer 3 (Behavioral guardrails, ~200 tokens, cached):** Specifies constraints such as prohibiting invented rates, assumptions of eligibility, and enforcing confidence thresholds (patches below 0.75 confidence are disallowed). Requires rule citation on every patch.
- **Layer 4 (Per-request context, 400–800 tokens, never cached):** Contains user intent, sheet name, column headers, selected rows as compact JSON, and prior session flags.

The caching of layers 1–3 allows for efficient reuse and cost savings in repeated requests.

### 10. Process: LLM_CALLER + RESPONSE_VALIDATOR

LLM_CALLER sends the assembled prompt to the chosen LLM and receives the JSON patch response. RESPONSE_VALIDATOR ensures the response adheres to the output schema and confidence requirements before passing data forward.

### 11. Process: CONTROLLED_REVEAL

CONTROLLED_REVEAL creates a smooth user experience by mimicking streaming behavior without actual streaming. The backend returns all patches in a single JSON payload; the frontend uses timed intervals (120ms) to reveal each diff card sequentially via `setTimeout`. This staged reveal gives the impression of incremental discovery.

### 12. Process: PATCH_WRITER

PATCH_WRITER executes immediately upon user acceptance of a suggested change. It uses Excel.run() to write the new value to the corresponding Excel cell, applies a brief green background flash (lasting 2 seconds for visual confirmation), and triggers AUDIT_LOGGER backend logging.

### 13. Caching Strategy

CELLIX employs three distinct caching layers targeting repeated computations at different stack levels, improving performance and reducing operational cost. These caches correspond to:

- Prompt components (GST rules, guardrails)
- Processed data segments
- Session and audit data retrieval

### 14. Process: AUDIT_LOGGER + MongoDB Schema

MongoDB is selected over PostgreSQL for audit logging due to its flexible, document-oriented schema, which naturally fits the variable patch data structure. Audit entries vary widely in fields depending on patch type, workflow, and source, making relational schema management cumbersome.

**Audit log features:**

- Append-only write model, well-suited for MongoDB
- Collections include: `audit_entries`, `sessions`, and `users`

#### 14.1 Why MongoDB over PostgreSQL

MongoDB’s flexibility accommodates varying audit entry structures without nullable fields or complex joins, facilitating efficient data storage and retrieval.

#### 14.2 MongoDB collections

- `audit_entries`: Stores detailed patch logs with timestamps, reasons, confidence scores, and user identifiers.
- `sessions`: Tracks session state and flags.
- `users`: Stores user metadata.

### 15. GST Reconciliation — Deep Dive

GST reconciliation is CELLIX's flagship workflow and core use case, representing the highest-value function. Understanding its complexities is vital for engineers.

#### 15.1 What GST reconciliation is

GST reconciliation involves comparing the taxpayer’s internal purchase register against the GST portal’s auto-populated GSTR-2A/2B data. The goal is to ensure that purchases recorded internally match what suppliers have reported to the government.

- GSTR-1: Supplier-filed sales returns
- GSTR-2A/2B: Auto-filled purchase data from suppliers

Mismatches block legitimate Input Tax Credit (ITC) claims, which represent real financial value and compliance risk.

#### 15.2 Why it is painful

- **Volume:** A CA managing 50 clients processes thousands of invoice rows monthly.
- **Data quality:** Supplier errors (wrong GSTIN, amounts, filing delays) cause mismatches.
- **Fuzzy matching:** Invoice numbers may vary in format between systems, e.g., “INV-101” vs. “inv101” vs. “101”.
- **Rate changes:** GST rates change periodically; correct rate application depends on invoice date.
- **ITC eligibility:** Certain categories are permanently blocked under Sec 17(5) CGST Act, which changes with amendments.

#### 15.3 The CELLIX GST reconciliation workflow

When INTENT_CLASSIFIER identifies a gst_recon request, the system runs RULES_ENGINE and LLM pipeline to automate matching, flagging, and patching discrepancies.

#### 15.4 Fuzzy matching logic

The RULES_ENGINE implements fuzzy matching algorithms to handle invoice number variations and inconsistencies, improving match accuracy while reducing manual review.

#### 15.5 What the CA sees — reconciliation output

The CA’s workbook displays:

- **Green cells:** Accepted patches (e.g., corrected rates, ITC flags)
- **Red/blue diff cells:** Pending changes with old value struck through in red and suggested value in blue
- **Amber flag cells:** Manual review items such as missing invoices or ambiguous HSN codes
- **Audit PDF:** A detailed ICAI-compliant working paper containing every change, timestamp, legal citation, CA email, and confidence score.

### 16. Process: PDF_EXPORTER

PDF_EXPORTER manages audit PDF generation asynchronously using BullMQ background jobs. When the CA requests an audit PDF export, the backend queues the job and returns a URL once the PDF is ready, with frontend polling every 2 seconds to update the user.

### 17. Full End-to-End Flow Summary

The document concludes with a comprehensive table listing every named process in execution order for a standard GST reconciliation request, providing a holistic view of CELLIX’s internal workflow from prompt capture through audit export.

---

This detailed architectural and process overview highlights CELLIX’s design strategies to optimize user experience, cost management, compliance accuracy, and audit traceability within the Indian GST compliance domain. The integration of deterministic rules with advanced LLM inference, combined with a sophisticated caching and audit strategy, enables scalable, reliable automation for complex spreadsheet-based workflows.