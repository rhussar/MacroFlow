# Task: Build Mode Default Macro Location (Selected Workbook -> New Module)

## Why
Build Mode should default to creating a new macro in the workbook currently selected in Search Mode (the workbook highlighted in the picker), not a hardcoded location.

## Current Gaps (review)
- Build Mode falls back to a hardcoded location: `PERSONAL.xlsm > Module2`.
  - `App/src/components/BuildMode.jsx:25`
  - `App/src/components/BuildMode.jsx:38`
- App does not pass workbook context into Build Mode.
  - `App/src/App.jsx:188`
  - `App/src/App.jsx:224`
- Search Mode owns workbook picker state locally, so selected workbook is not shared with Build Mode.
  - `App/src/components/SearchMode.jsx:43`
  - `App/src/components/SearchMode.jsx:47`
  - `App/src/components/SearchMode.jsx:174`
- Build Mode entry button sends no workbook/module payload.
  - `App/src/components/SearchMode.jsx:428`
- Backend inject API is active-workbook scoped only; no workbook-targeted inject endpoint yet.
  - `App/electron/ipc-handlers.js:484`

## Goal
When user enters AI Build Mode, default target is:
1. Workbook: the currently selected workbook from Search Mode.
2. Module: a new standard module in that workbook (auto-generated unique name).
3. UI: the breadcrumb above code shows that workbook and pending new module destination.

## Scope
- Include Build Mode entry from:
  - AI Build button (`SearchMode` header)
  - Keyboard shortcut (`Alt+M`)
- Keep behavior deterministic and scalable across many open workbooks.
- Do not include AI prompt/codegen quality changes in this task.

## Functional Requirements
- Use selected workbook (picker selection), not always active workbook.
- If selected workbook is missing/closed at submit time, show actionable error and allow retry/retarget.
- Default module strategy is `create_new`.
- Module naming must avoid collisions (e.g. `Module1`, `Module2`, ... or `MacroFlowModule1`, ...).
- Breadcrumb always reflects real target: `<Workbook> > <ModuleName (new)>`.
- After successful save, module appears in All Files/Explorer without app restart.

## Technical Design
### 1) Share workbook target state across modes
- Lift selected workbook context to `App` state (or a shared hook owned by `App`) and pass down:
  - `selectedWorkbookForBuild`
  - `onSelectedWorkbookChange`
- `SearchMode` keeps UI picker but writes selected workbook to this shared state.
- `BuildMode` reads this shared state as default target.

### 2) Introduce Build target contract
Create a stable model used by UI + IPC:
```ts
BuildTarget = {
  workbook: { name: string, path: string, key: string },
  module: {
    mode: 'create_new' | 'existing',
    name: string,
  }
}
```
- Default `mode` is `create_new`.
- Resolve module name before inject.

### 3) Backend workbook-scoped inject
Add IPC + bridge methods to inject into a specific workbook (do not depend on current active workbook):
- New channel: `vba:inject:by-workbook`
- Request: `{ workbookName, workbookPath, moduleName, code, createIfMissing?: boolean }`
- Behavior:
  - Find target workbook by path/name (reuse existing workbook resolution pattern from by-workbook handlers).
  - Create module when missing and `createIfMissing` true.
  - Return `{ success, workbookFound, workbook, moduleName, message }`.

### 4) Build Mode submit flow
- On first submit in empty state:
  - Resolve workbook target from shared selection.
  - Fetch workbook modules (`vba.modulesByWorkbook`) to compute next unique module name.
  - Update breadcrumb immediately.
  - Save generated code via workbook-scoped inject.
- Keep code preview editable as today.

### 5) Refresh/consistency
- After successful save, trigger search/workbook refresh path so module list updates.
- Ensure explorer context can deep-link to newly created module later.

## Acceptance Criteria
- Clicking AI Build Mode with workbook `ClientA.xlsm` selected creates a new module in `ClientA.xlsm` by default.
- Alt+M uses the same selected workbook default.
- Breadcrumb displays selected workbook + new module name before/at save time.
- If selected workbook closes before save, user sees clear error and no silent fallback to active workbook.
- No hardcoded `PERSONAL.xlsm > Module2` fallback remains in runtime behavior.
- Existing run/shortcut/search behavior remains unchanged.

## Test Plan
### Unit
- `useWorkbookPickerData` integration point: selected workbook propagation to App state.
- New module name generator handles collisions and sparse sequences.
- Build target reducer/state transitions.

### Electron/IPC
- `vba:inject:by-workbook` success path.
- `vba:inject:by-workbook` workbook-not-found path.
- Module created when missing and code inserted correctly.

### UI/Integration
- Enter Build Mode from Search button and Alt+M with non-active workbook selected.
- Breadcrumb correctness before and after save.
- Workbook closed during flow -> expected error state.

## Implementation Steps
1. Lift workbook selection state from SearchMode to App-level shared state.
2. Wire `onBuildModeClick` to pass selected workbook context into BuildMode.
3. Replace hardcoded BuildMode default location with BuildTarget state.
4. Add workbook-scoped inject API (preload + ipc handler + excel bridge).
5. Add unique module name resolver and submit-time target resolution.
6. Refresh workbook/module data after save.
7. Add/extend tests for state, IPC, and UI flows.

## Out of Scope
- Prompt engineering / code generation quality.
- Full Build Mode UX redesign.
- Cross-workbook bulk operations.
