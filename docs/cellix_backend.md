### 1. Why NestJS for CELLIX Backend

The CELLIX backend is a stateful, multi-module NestJS service covering LLM orchestration, conversation management, deterministic routing, audit logging, and change sets. See [`docs/CELLIX_TECHNICAL_DOCUMENTATION.md`](../../docs/CELLIX_TECHNICAL_DOCUMENTATION.md) for the current module map and setup.

> **Note:** Sections below describe a planned GST/rules-engine architecture (Redis, BullMQ, JWT auth) that is **not** implemented in Cellix-2026. Use the main technical doc for accurate setup instructions.

### 2. Project Initialisation

#### 2.1 Prerequisites

To start the CELLIX backend project, the following tools and environments are necessary:

- **Node.js 20 LTS:** Verify installation via `$ node --version $`. Use Node Version Manager (nvm) to install if missing.
- **pnpm 9.x:** Preferred over npm due to its monorepo-friendly installation and faster dependency management. Install globally using `$ npm install -g pnpm $`.
- **NestJS CLI:** Required to scaffold and manage the project, installed globally via `$ npm install -g @nestjs/cli $`.
- **Docker Desktop:** Used to run local MongoDB and Redis instances via containers, avoiding direct installs on the development machine.
- **MongoDB Atlas Account:** For development, a free tier cluster suffices, while production requires a dedicated cluster.

#### 2.2 Create the NestJS Project

The NestJS CLI initializes the project structure, setting the foundation for modular development.

#### 2.3 Install all Dependencies Upfront

All required dependencies, including NestJS packages, MongoDB drivers, Redis clients, and LLM libraries, are installed at once to ensure consistency and reduce version conflicts.

#### 2.4 Switch to Fastify (Performance)

By default, NestJS uses Express as the HTTP adapter. For CELLIX, Fastify is preferred because it delivers 2–3 times faster performance for lightweight JSON APIs typical of CELLIX’s workload. The switch is done early in the `main.ts` file, before building other parts of the application.

### 3. Project Structure — Full Module Map

The project follows a strict modular organization. Each feature or domain capability resides in its own directory under `src/`. No business logic resides in the root `app.module.ts`; instead, it serves solely as a wiring module importing all feature modules. This clear separation enhances maintainability and enforces boundaries between features.

### 4. Config Module

The configuration module wraps the `@nestjs/config` package and integrates a Joi validation schema. This validation enforces the presence and correctness of environment variables, such as `$ ANTHROPIC_API_KEY $`. The key benefit is fail-fast behavior during application bootstrap, avoiding runtime errors when environment variables are missing or invalid.

#### 4.1 Environment Variables

Critical environment variables include API keys, database URIs, Redis host details, and JWT secrets. These must be defined in the `.env` file.

#### 4.2 Config Schema Validation

Joi schemas validate environment variables, enforcing types and presence constraints. This prevents misconfiguration from causing subtle bugs.

#### 4.3 Config Module and Typed Service

The module exposes a strongly typed service interface to access environment variables, ensuring type safety across the codebase.

### 5. Database Module — MongoDB + Mongoose

Each module owning a MongoDB collection defines its own Mongoose schema in a dedicated `schemas/` directory. Schemas are encapsulated within their respective modules to avoid cross-imports and tightly coupled data models, promoting clear ownership and maintainability.

### 6. Cache Module — Redis

The caching layer uses Redis to improve performance and enable features like caching LLM responses or session data. Redis is run locally using Docker during development.

### 7. Auth Module

Authentication is handled via JWTs, with the following components:

#### 7.1 JWT Strategy

A JWT strategy validates tokens, ensuring secure access to protected endpoints.

#### 7.2 Credits Guard

This guard enforces usage limits or credits, potentially linked to API usage or feature access control.

#### 7.3 Auth Module Wiring

The module combines strategies and guards, exposing authentication endpoints and middleware.

### 8. Analyse Module — Core Request Handler

This module is the backbone of the backend pipeline. It receives POST requests from the Excel task pane and orchestrates all submodules sequentially:

- **ClassifierModule**
- **RulesEngineModule**
- **LlmModule**
- **AuditModule**

#### 8.1 Request DTO

Defines the data structure expected in requests, ensuring validation and consistency.

#### 8.2 Controller

Handles HTTP requests, invoking the service layer.

#### 8.3 Service — Pipeline Orchestration

Coordinates the execution flow of classification, rules evaluation, LLM calls, and auditing in a strict sequence to produce deterministic results.

### 9. Classifier Module

This module categorizes or classifies incoming data, acting as the first step in the pipeline. It helps route requests or determine processing logic downstream.

### 10. Rules Engine Module

#### 10.1 Module Structure

Organized to encapsulate all business rules logic.

#### 10.2 GstRate Schema (MongoDB)

Defines MongoDB schema for GST rates, which are critical for tax calculations and validations.

#### 10.3 Rate Validator Service

Validates rates against business rules, ensuring correctness before further processing.

### 11. LLM Module

Handles communication with large language models.

#### 11.1 LLM Module Wiring

Wires dependencies and services necessary for LLM operations.

#### 11.2 Model Selector Service

Chooses the appropriate LLM model based on request context or parameters.

#### 11.3 LLM Service — Anthropic Call with Caching

Executes calls to the Anthropic API, integrating caching to minimize redundant requests and improve response times.

### 12. PDF Module — BullMQ Background Job

Handles PDF generation and processing asynchronously using BullMQ, a Redis-based queue system. This offloads heavy or time-consuming jobs from the main request-response cycle, improving responsiveness.

### 13. Root App Module — Wiring Everything Together

The root module imports all feature modules but contains no business logic itself. The import order is significant, especially ensuring that `BullModule` is imported after the config module to access environment variables correctly.

### 14. Docker Compose — Local Development

Local MongoDB and Redis instances run via Docker Compose, avoiding direct installation on developers' machines. This facilitates easy teardown, reset, and environment replication.

#### 14.1 MongoDB Seed Script

The `gst_rates` collection must be seeded for the rules engine to function. The seed script (`scripts/mongo-seed.js`) runs automatically on the first Docker startup, ensuring the database is pre-populated.

### 15. Testing

#### 15.1 Unit Testing with NestJS Test Module

Unit tests leverage NestJS’s testing utilities to isolate modules and verify individual components.

#### 15.2 E2E Test

End-to-end tests validate entire request pipelines, ensuring all modules integrate correctly.

#### 15.3 Running Tests

Tests are run via `$ pnpm test $`, and they depend on the seeded local MongoDB instance.

### 16. Common Decorators and Utilities

#### 16.1 CurrentUser Decorator

Simplifies access to the authenticated user in controllers.

#### 16.2 Global Exception Filter

Centralizes error handling, formatting exceptions uniformly across the API.

### 17. Scripts and Useful Commands

#### 17.1 Generating New Modules with NestJS CLI

Developers are encouraged to use NestJS CLI to scaffold new modules, ensuring consistent file structure and wiring, reducing human error.

### 18. Module Dependency Graph

The dependency graph shows module imports and ensures no circular dependencies exist. The `AnalyseModule` acts as a leaf node, consuming services from all other feature modules without being imported elsewhere.

### 19. Quick-Start Checklist

A step-by-step setup guide for new developers:

1. Clone the repository and install dependencies:  
   $$ \texttt{pnpm install} $$
2. Copy `.env.example` to `.env` and fill environment variables such as MongoDB URI, Redis host, Anthropic API key, and JWT secret.
3. Start local databases using Docker Compose:  
   $$ \texttt{docker-compose up -d} $$
4. Verify seeding of GST rates:  
   $$ \texttt{mongosh mongodb://localhost:27017/cellix --eval "db.gst_rates.countDocuments()"} $$  
   The count should be greater than zero.
5. Start the development server:  
   $$ \texttt{pnpm start:dev} $$  
   The server should boot without errors and listen at http://localhost:3000.
6. Visit Swagger documentation at http://localhost:3000/api/docs to confirm all routes are registered.
7. Run unit tests:  
   $$ \texttt{pnpm test} $$  
   All tests should pass against the seeded database.
8. Make a test POST request to `/api/analyse` with a valid JWT obtained from `/api/auth/login` to verify the pipeline runs end-to-end.

---

This guide ensures a robust, scalable backend setup leveraging NestJS’s modular architecture, MongoDB for persistence, Redis for caching and queues, and integration with LLM APIs, all orchestrated in a clean, maintainable codebase ready for production deployment.