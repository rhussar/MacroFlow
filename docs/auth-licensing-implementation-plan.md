# Auth, SSO, And Licensing Implementation Plan

## Purpose

This document is an implementation plan for adding:

- user sign-in and sign-out
- team and organization support
- optional enterprise SSO
- desktop licensing and device activation
- future offline licensing support

It is written for the current MacroFlow app, which is an Electron desktop application under `App/`. As of 2026-03-15, the repo does not contain existing Auth0 or Keygen integration code.

## Short Recommendation

Use:

- Auth0 for authentication, users, and organizations
- Keygen.sh for license issuance, activations, and entitlements
- a small backend you control to connect the two

Do not build authentication or licensing primitives from scratch.

Do not make the Electron app talk directly to all vendor APIs with no backend.

## Recommended V1 Scope

Ship these first:

- Auth0 native-app login using Authorization Code + PKCE
- sign-in and sign-out in Electron
- token refresh and local session restore
- user profile and organization context
- Keygen node-locked licenses
- device activation for the current machine
- entitlement checks for gated features
- short offline grace window, not true air-gapped offline
- small admin flow for assigning and revoking seats

Defer these until later:

- enterprise SAML / OIDC connections
- SCIM provisioning
- full air-gapped offline licensing
- self-serve billing automation if you want to move faster

## Why This Architecture

Auth0 is the right place for:

- login UX
- identity providers
- MFA
- user accounts
- organizations / tenants
- future enterprise SSO

Keygen is the right place for:

- license keys
- machine activation
- device limits
- entitlement flags
- offline license files later

Your backend is still required for:

- mapping Auth0 users and orgs to your product records
- assigning licenses
- enforcing product rules
- handling billing webhooks
- storing audit history
- shielding admin/vendor secrets from the desktop app

## Target Architecture

### Components

1. Electron app
   Responsibilities:
   - launch sign-in
   - store session securely
   - call backend APIs
   - show license state
   - activate and deactivate current device
   - gate features based on entitlements

2. Auth backend
   Responsibilities:
   - verify Auth0 tokens
   - create/update local user records
   - map user to organization and role
   - mint a short-lived app session if desired

3. Licensing backend
   Responsibilities:
   - create and assign Keygen licenses
   - activate current machine
   - check license status
   - store license and activation history
   - expose entitlement state to the app

4. Postgres database
   Responsibilities:
   - users
   - organizations
   - memberships
   - licenses
   - device activations
   - audit logs
   - subscription state

5. Auth0
   Responsibilities:
   - identity
   - login UX
   - social login
   - future enterprise SSO
   - future organizations

6. Keygen
   Responsibilities:
   - license resources
   - policies
   - machines
   - entitlements
   - later offline license files

7. Optional billing provider
   Recommended later:
   - Stripe

## Recommended Technology Choices

Use a separate backend service, not Electron main process code, for vendor orchestration.

Recommended stack:

- Backend: Node.js + TypeScript + Fastify
- Database: Postgres
- ORM: Prisma or Drizzle
- Deployment: Render, Fly.io, Railway, or a small VPS
- Secrets: deployment platform secrets manager
- Electron secure storage: OS keychain via `keytar`, or Windows Credential Manager equivalent

If you want one repo:

- keep `App/` for Electron
- add `Server/` for the backend

If you want cleaner ownership:

- use a second repo for the backend

## Product Decisions To Lock Before Implementation

These decisions must be made before the students start writing code:

1. Account model
   Choose one:
   - individuals only
   - teams only
   - both

2. License model
   Choose one:
   - per-user seat
   - per-device seat
   - per-team shared seat

3. Device policy
   Example:
   - 2 devices per paid seat

4. Offline policy
   Choose one:
   - online only
   - grace window only
   - true offline / air-gapped

5. Provisioning model
   Choose one:
   - manual admin assignment first
   - automatic Stripe provisioning in V1

6. Enterprise SSO timing
   Choose one:
   - not in V1
   - support it in V1 but hide behind plan gating

## Recommended V1 Product Decisions

To keep scope sane for two developers:

- account model: both individuals and teams
- license model: per-user seat with up to 2 machines
- offline policy: 7-day grace period only
- provisioning: manual admin assignment first
- enterprise SSO: not in V1

That gives you a solid commercial desktop app without taking on the hardest identity and offline edge cases immediately.

## End-To-End Flows

### 1. First-Time Sign-In

1. User clicks `Sign in` in Electron.
2. Electron opens the system browser, not an embedded webview.
3. User authenticates with Auth0 Universal Login.
4. Auth0 redirects back to the desktop app using a native-app OAuth callback.
5. App exchanges the code with PKCE and gets tokens.
6. App stores refresh/session material in the OS keychain.
7. App calls your backend `POST /auth/session/exchange`.
8. Backend verifies the Auth0 token.
9. Backend upserts local user and membership records.
10. Backend returns the app session plus current entitlements and license summary.
11. App routes to the signed-in state.

### 2. App Startup With Existing Session

1. App reads secure local auth state.
2. If refresh is valid, it refreshes tokens.
3. App calls `GET /me` and `GET /license/status`.
4. If license is valid, app unlocks entitled features.
5. If login is valid but license is missing, app shows `No active license`.

### 3. Seat Assignment

1. Admin purchases or is manually assigned seats.
2. Backend creates or updates a local subscription record.
3. Backend creates a Keygen license or attaches the user to an existing org license.
4. Backend stores the Keygen license ID and policy.
5. User’s next session refresh returns active entitlements.

### 4. Device Activation

1. App computes a stable device fingerprint.
2. App calls backend `POST /license/activate-device`.
3. Backend validates that the signed-in user is entitled to activate.
4. Backend validates license state with Keygen.
5. Backend creates or reuses a Keygen machine activation.
6. Backend stores the activation locally.
7. App persists the activation record locally and continues.

### 5. Sign-Out

1. App clears its local session and tokens.
2. App does not automatically deactivate the machine on every logout.
3. App offers a separate `Deactivate this device` action.

This avoids accidental seat churn when users sign in and out.

### 6. Offline Grace Window

1. On successful online validation, backend returns:
   - entitlement state
   - license expiry/check-in deadline
   - grace window deadline
2. App stores this signed snapshot locally.
3. If the app is offline but the grace window is still valid, app continues with limited checks.
4. If the grace window expires, app requires revalidation.

### 7. Future True Offline / Air-Gapped Flow

This is a V2 feature:

1. Backend or admin checks out a signed Keygen license file.
2. User imports `.lic` file into the desktop app.
3. App verifies the license file cryptographically using the Keygen public key.
4. App respects the file TTL and requires re-checkout when expired.

## Auth0 Setup

Create:

- one Auth0 tenant
- one native application for Electron
- one API for your backend

Configure:

- Universal Login
- Authorization Code + PKCE
- refresh token rotation
- social login providers if needed
- Organizations only if you want teams in V1

Do not:

- embed a client secret in the Electron app
- rely on an embedded browser login flow

### Auth0 Notes For Electron

- Native apps should use PKCE because they cannot safely store a client secret.
- Auth0 warns against custom URI schemes because malicious apps can potentially intercept them.
- If you need team accounts, Auth0 Organizations supports Individuals, Business Users, or Both.

## Keygen Setup

Create:

- one Keygen product
- at least one policy
- one environment for production
- one environment for staging

Recommended V1 policy:

- node-locked
- max 2 machines
- expiration tied to subscription period
- required entitlements for premium features

Start with:

- one plan/policy for free trial
- one plan/policy for paid Pro
- optional team policy later

### Keygen Notes

- Machine activation is a client-driven flow, but admin/product tokens must not be embedded in the app.
- Use backend-issued actions to create or revoke machines.
- Entitlements should drive feature gating.
- For future offline mode, use checked-out license files with TTL.

## Local Data Model

Recommended tables:

### `users`

- `id`
- `auth0_user_id`
- `email`
- `name`
- `avatar_url`
- `created_at`
- `updated_at`

### `organizations`

- `id`
- `auth0_org_id`
- `name`
- `slug`
- `billing_email`
- `created_at`
- `updated_at`

### `memberships`

- `id`
- `user_id`
- `organization_id`
- `role`
- `status`
- `created_at`
- `updated_at`

### `licenses`

- `id`
- `organization_id` nullable
- `user_id` nullable
- `keygen_license_id`
- `policy_code`
- `status`
- `seat_count`
- `max_machines_per_seat`
- `expires_at`
- `last_validated_at`
- `created_at`
- `updated_at`

### `license_entitlements`

- `id`
- `license_id`
- `code`
- `value`
- `created_at`
- `updated_at`

### `device_activations`

- `id`
- `license_id`
- `user_id`
- `keygen_machine_id`
- `fingerprint_hash`
- `device_name`
- `platform`
- `status`
- `last_heartbeat_at`
- `created_at`
- `updated_at`

### `audit_logs`

- `id`
- `actor_user_id`
- `organization_id`
- `event_type`
- `payload_json`
- `created_at`

### `subscriptions`

- `id`
- `organization_id` nullable
- `user_id` nullable
- `provider`
- `provider_subscription_id`
- `status`
- `plan_code`
- `renewal_at`
- `created_at`
- `updated_at`

## Backend API Surface

Minimum endpoints:

### Auth

- `POST /auth/session/exchange`
- `POST /auth/refresh`
- `POST /auth/logout`
- `GET /me`

### Licensing

- `GET /license/status`
- `POST /license/activate-device`
- `POST /license/deactivate-device`
- `POST /license/check-in`
- `GET /license/entitlements`

### Admin

- `POST /admin/licenses/assign`
- `POST /admin/licenses/revoke`
- `POST /admin/organizations/invite`
- `GET /admin/organizations/:id/users`
- `GET /admin/licenses/:id/devices`

### Webhooks

- `POST /webhooks/stripe`
- `POST /webhooks/auth0`
- `POST /webhooks/keygen`

## Electron App Changes

### New UI Areas

- Settings -> Account
- Settings -> License
- Sign-in screen
- Signed-out screen
- License expired / missing state
- Device management modal

### New Electron/Main Responsibilities

- open system browser for login
- handle OAuth callback
- securely store tokens
- expose session info through preload APIs
- expose secure logout

### New Renderer Responsibilities

- react auth state
- restore session on boot
- show org picker if user belongs to multiple orgs
- show license state
- gate AI or premium features

### Recommended Client Modules

Under `App/electron/`:

- `auth0-client.js`
- `session-store.js`
- `license-client.js`
- `device-fingerprint.js`

Under `App/src/features/`:

- `auth/`
- `licensing/`
- `account/`

## Student Work Split

## Student A: Electron + Auth

Primary ownership:

- Auth0 desktop login flow
- secure token storage
- session restore
- sign-in / sign-out UI
- account settings UI
- app state gating by auth state
- integration tests for Electron auth/session behavior

Deliverables:

- working login/logout flow
- persistent session restore
- signed-out and signed-in app states
- secure local token handling

## Student B: Backend + Licensing

Primary ownership:

- backend service scaffold
- database schema
- Auth0 token verification
- Keygen integration
- license status APIs
- device activation APIs
- admin assignment flows
- webhook handling

Deliverables:

- deployable backend
- seeded staging environment
- license assignment and activation working
- entitlement API working
- admin/debug tooling for support

## Shared Responsibilities

- define product roles and entitlements
- align auth state and license state contracts
- test org switching and seat revocation flows
- document operational runbooks

## Milestone Plan

## Milestone 0: Decisions And Setup

Time:

- 2 to 3 days

Tasks:

- lock product decisions from the earlier section
- create Auth0 tenant
- create Keygen account/product/policies
- choose backend hosting
- choose database provider
- create staging secrets

Exit criteria:

- all secrets provisioned
- staging Auth0 and Keygen environments exist
- callback URLs and allowed origins are agreed

## Milestone 1: Backend Skeleton And Local Data Model

Time:

- 4 to 5 days

Tasks:

- create backend project
- add database and migrations
- add health endpoint
- add auth middleware
- create local tables
- add structured logging

Exit criteria:

- deployable staging backend
- database schema migrated
- authenticated `/me` endpoint works with mocked token

## Milestone 2: Electron Authentication

Time:

- 4 to 6 days

Tasks:

- implement PKCE login
- handle callback
- add token storage
- restore session on app boot
- implement logout
- add account UI state

Exit criteria:

- sign in works in packaged Electron build
- sign out works
- session persists across restart

## Milestone 3: License Assignment And Status

Time:

- 4 to 6 days

Tasks:

- implement Keygen service wrapper
- create license mapping tables
- implement `/license/status`
- implement manual seat assignment
- add entitlement contract

Exit criteria:

- signed-in user can be assigned a license
- app receives entitlements correctly
- unlicensed users are gated correctly

## Milestone 4: Device Activation

Time:

- 4 to 5 days

Tasks:

- implement fingerprint generation
- implement activate/deactivate APIs
- create machine records in Keygen
- handle duplicate activations
- add device management UI

Exit criteria:

- user can activate current machine
- device limit is enforced
- deactivation works

## Milestone 5: Grace Window And Hardening

Time:

- 4 to 5 days

Tasks:

- add cached license snapshot
- add grace window logic
- handle expired sessions
- add audit logging
- add rate limiting and defensive checks
- test seat revocation and subscription expiry

Exit criteria:

- app handles temporary offline use gracefully
- revoked or expired users are blocked correctly after grace expiry

## Milestone 6: Admin And Operations

Time:

- 3 to 5 days

Tasks:

- add minimal admin pages or admin scripts
- add support playbook
- add webhook replay procedure
- add staging test checklist

Exit criteria:

- support can assign and revoke seats
- support can inspect device activations
- webhook failures are diagnosable

## Total Estimate

For two capable master's students:

- focused full-time: 4 to 6 weeks
- part-time: 6 to 10 weeks

This assumes:

- no enterprise SSO in V1
- no true offline licensing in V1
- no full self-serve billing portal in V1

## Acceptance Criteria

The project is complete when all of these are true:

- a new user can sign in from a packaged desktop build
- the app restores the session after restart
- the backend recognizes the signed-in user and org
- the user can be assigned a license
- the current device can be activated
- premium features are gated by entitlement
- sign-out fully clears local auth state
- revoked or expired licenses stop working after the grace window
- staging has a documented support/admin path

## Key Risks

1. Electron OAuth callback handling
   Mitigation:
   - test early in packaged builds, not only dev mode

2. Device fingerprint instability
   Mitigation:
   - hash a stable hardware identifier and test upgrades/reinstalls

3. Seat churn from logout
   Mitigation:
   - separate logout from device deactivation

4. Over-scoping enterprise SSO
   Mitigation:
   - do not include SAML/SCIM in V1 unless a customer already requires it

5. Over-scoping full offline licensing
   Mitigation:
   - ship grace-window offline first

6. Desktop secrets leakage
   Mitigation:
   - do not embed vendor admin secrets in Electron

## Suggested Testing Plan

### Unit Tests

- token parsing
- entitlement mapping
- fingerprint normalization
- grace window logic

### Integration Tests

- sign-in session exchange
- sign-out
- license assignment
- activation and deactivation
- org switching

### Manual QA

- packaged app login on Windows
- lost network during boot
- expired refresh token
- revoked seat
- second device activation hitting limit

## What I Would Not Put In V1

- SCIM
- domain discovery
- enterprise IdP setup UI
- true air-gapped offline
- floating seat pools
- usage-based billing
- customer self-serve admin portal

These are good V2 items, but they will slow down a clean first release.

## Implementation Order I Recommend

1. build the backend skeleton
2. implement Auth0 login in Electron
3. wire `/me` and session restore
4. implement license status
5. implement device activation
6. add entitlement gating in the app
7. add admin assignment flow
8. add grace window
9. add operational tooling

## Sources

- Auth0 Authorization Code Flow with PKCE:
  https://auth0.com/docs/get-started/authentication-and-authorization-flow/authorization-code-flow-with-pkce
- Auth0 Organizations login flows:
  https://auth0.com/docs/manage-users/organizations/login-flows-for-organizations
- Auth0 Refresh Token Rotation:
  https://auth0.com/docs/secure/tokens/refresh-tokens/refresh-token-rotation
- Keygen machine activation:
  https://keygen.sh/docs/activating-machines/
- Keygen entitlements:
  https://keygen.sh/docs/api/entitlements/
- Keygen offline licensing and cryptographic verification:
  https://keygen.sh/docs/api/cryptography/
- Keygen license checkout for offline and air-gapped usage:
  https://keygen.sh/docs/api/licenses/
