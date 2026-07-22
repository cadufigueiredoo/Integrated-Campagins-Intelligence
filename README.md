# Integrated Campaign Intelligence

Plan, localize, measure, and optimize integrated B2B campaigns in one place.

**Developed by Carlos Eduardo** - [linkedin.com/in/carloseduardovf](https://www.linkedin.com/in/carloseduardovf/)

---

## What it does

A campaign is a single object that moves through four phases:

| Phase | What happens |
|---|---|
| **Plan** | Thesis and timing, personas by funnel stage, multichannel mix, message architecture, asset plan, and user-defined KPI targets. |
| **Localize** | Regional adaptation: language, tone, local channels, LGPD compliance, cultural notes, local calendar. |
| **Measure** | Upload the campaign's spreadsheet extract. The app computes CTR, CPL, CPO, stage conversion, ROAS and pipeline, then charts them. |
| **Optimize** | Attainment against your targets, plus an executive summary and recommendations ranked by priority. |

## Design principle: data integrity

**The app computes every number. The AI only writes.**

- Every metric (CPL, CTR, ROAS, conversion rates, attainment) is calculated in JavaScript from the uploaded data.
- The model receives already-computed figures and narrates them. It never produces a metric.
- KPI targets are entered by the user, never generated. The benchmark is always human-owned.
- Attainment inverts for cost metrics, so above 100% always means outperforming, whichever metric you look at.

## Bilingual by design

The interface is English. Generated content follows the **Output language** toggle (PT/EN) per campaign, so the same profile serves the Brazilian market or a global stakeholder. Switching language on an existing campaign translates the content in place; funnel stages, metric identifiers, proper nouns and all numbers are preserved.

The bundled sample campaign ships with both language versions, so it switches instantly with no network call.

## Data input contract

Measure expects a standard column set. Download the template from inside the app, or match these headers:

```
Channel | Funnel Stage | Spend | Impressions | Clicks | Leads | MQL | SQL | Opportunities | Pipeline | Revenue
```

Column matching tolerates naming variations in Portuguese and English, and the number parser handles Brazilian formatting (R$ 1.710.000,00) as well as US formatting. Funnel stage labels such as topo, TOFU, meio, MOFU, fundo and BOFU are normalized automatically.

---

## Local development

```bash
npm install
npm run dev
```

Note: `/api/generate` only runs on Vercel's runtime. To exercise generation locally, use `vercel dev` instead of `npm run dev`.

## Deployment (Vercel)

1. Push this repository to GitHub, preserving the folder structure below.
2. In Vercel, import the repository. The framework preset should be detected as **Vite**.
3. Add the environment variable **before** the first deploy:
   - `ANTHROPIC_API_KEY` = your key from console.anthropic.com
4. Deploy.

`vercel.json` sets `maxDuration: 60` for the serverless function, which prevents timeouts on longer generations. The free Hobby tier covers this.

## Structure

```
index.html
vercel.json          Serverless function timeout
package.json
vite.config.js
.gitignore
README.md
src/main.jsx         React entry point
src/App.jsx          Entire application (single file)
api/generate.js      Serverless function; holds ANTHROPIC_API_KEY
```

- **Frontend:** React + Vite, Recharts for charts, SheetJS for spreadsheet parsing, lucide-react for icons.
- **Persistence:** campaigns are saved in `localStorage` under the `ici:campaign:` prefix. No database, no login.
- **Resilience:** generation is split into small batches so responses fit the output budget; truncated JSON is repaired rather than discarded; the serverless function retries with exponential backoff on 429/529/5xx and maps credit-balance and auth failures to clear messages.
- **Robustness:** campaigns saved under older schemas are normalized on load, so a missing field never breaks a phase.

## Roadmap (v2)

- Contact and account enrichment via Lusha / Apollo APIs
- Native PPTX export of the results dashboard
- RACI matrix, launch timeline, and risk register
- Multiple saved company profiles
- Server-side persistence with authentication
