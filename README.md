# Training Logger

Personal workout logger. Static PWA (`index.html`, GitHub Pages) + Google Apps Script web app (`Code.gs`) bound to a Google Sheet. The sheet is the source of truth; the app syncs profile, preferences, names, archive, templates and history across devices.

## Setup
1. Paste `Code.gs` into the sheet's Apps Script project, keep your own `SECRET`.
2. Deploy > Manage deployments > pencil > New version (Execute as me, access Anyone).
3. Open the Pages site, enter the web app URL + secret in the app.

## Tabs
| Tab | Columns |
|---|---|
| Sets | A Timestamp, B Day, C Exercise, D Set no., E Weight, F Reps target, G Reps done, H–K extras, L ID |
| Workouts | ID, Date, Workout, Start, End, Duration (min), Total volume, Sets, PRs |
| Template | Day, Order, Exercise, Sets, Rep low, Rep high, Rest, **Data** (JSON: e.g. `{"ss":true}` superset link, `{"rest":90}`, instructions) |
| Targets | Written each weekday morning by the scheduled task; read by the app |
| Rules | Progression rules used by the targets job |
| Records | Per-exercise records |
| Body | A–M fixed metrics, N **Data (JSON)** for extra metrics. Created on first body write |
| Settings | Key, Value (JSON), UpdatedAt. Last write wins by UpdatedAt |
| Exercises | Name, Body part, Category, Notes, Data (JSON), UpdatedAt |

## JSON columns
Rows may carry JSON that the client parses (relations such as supersets, per-exercise rest, extra metrics) so the sheet stays flat while objects are still expressible.

## Planning with Claude
A Claude chat with the Google Sheets connector can edit Template, Targets and Settings directly; the app picks the changes up on next sync/refresh.
