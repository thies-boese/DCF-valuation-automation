<div align="center">
  <h1>📊 Automated DCF Valuation Engine</h1>
  <p>
    <b>Enterprise-Grade Discounted Cash Flow Automation</b><br />
    <i>Transforming raw financial data into structured intrinsic valuation outputs.</i>
  </p>

https://github.com/user-attachments/assets/fbe8f50b-465c-44ec-a6a9-d45d844f194f



  <p>
    <img src="https://img.shields.io/badge/Node.js-339933?style=for-the-badge&logo=node.js&logoColor=white" alt="Node.js" />
    <img src="https://img.shields.io/badge/Excel%20Automation-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white" alt="Excel Automation" />
    <img src="https://img.shields.io/badge/Valuation-111827?style=for-the-badge" alt="Valuation" />
    <img src="https://img.shields.io/badge/DCF-0F766E?style=for-the-badge" alt="DCF" />
  </p>

  <p>
    <b>Audience:</b> Investment Banking recruiters &amp; MSc Finance admissions committees<br />
    <b>Focus:</b> Analyst-grade valuation workflow standardization, transparency, and reproducibility
  </p>
</div>

<hr />

<h2>🧠 Project Overview</h2>
<p>
  This project is a fully automated <b>intrinsic valuation engine</b> designed to mirror the workflow of an
  Investment Banking / Corporate Finance analyst. It retrieves structured company data, normalizes financial statement items,
  constructs a <b>FCFF-based DCF model</b>, computes <b>WACC</b>, forecasts operating performance using structured drivers,
  calculates terminal value via multiple methodologies, and exports a standardized Excel output for auditability and repeatability.
</p>

<p>
  The engine is built in <b>Node.js</b> and is optimized around one goal: <b>turn valuation theory into a deterministic system</b>
  while preserving analytical integrity (sanity checks, fallbacks, constraints).
</p>

<hr />

<h2>🎯 Strategic Purpose</h2>
<p>
  The project answers a professional question:
</p>
<blockquote>
  <b>Can valuation workflows be standardized and automated without losing analytical rigor?</b>
</blockquote>
<p>
  Rather than rebuilding spreadsheets manually for every company, the engine creates a consistent valuation pipeline:
  <b>data → normalization → FCFF → discounting → terminal value → enterprise → equity → standardized output</b>.
  This is aligned with real-world analyst workflows in pitchbooks, screening, and internal valuation memos.
</p>

<hr />

<h2>⚙️ End-to-End Workflow (What the Engine Does)</h2>

<h3>1) Company Selection &amp; Data Retrieval</h3>
<p>
  User inputs a ticker symbol. The backend fetches structured company data (financial statements + market data) and validates:
  missing values, null fields, inconsistent currencies, and unstable denominators.
</p>
<ul>
  <li><b>Objective:</b> prevent “garbage in → garbage out” modeling issues.</li>
  <li><b>Output:</b> canonical financial data object used downstream.</li>
</ul>

<h3>2) Financial Statement Normalization</h3>
<p>
  The engine standardizes line items into a consistent schema and separates:
  <b>operating items</b> vs <b>financing items</b> vs <b>non-recurring adjustments</b>.
  This ensures clean mapping into valuation mechanics, replicating analyst discipline.
</p>

<h3>3) FCFF Construction (Core Valuation Engine)</h3>
<p>
  Free Cash Flow to Firm is computed deterministically:
</p>
<pre><code>FCFF = EBIT(1 - Tax Rate) + D&amp;A - CapEx - ΔWorking Capital</code></pre>
<p>
  Each component is derived from the normalized dataset. Where data is incomplete, the engine applies transparent fallback logic
  and logs assumptions to preserve interpretability.
</p>

<h3>4) Forecasting Logic</h3>
<p>
  Forecasts are driven by structured assumptions (drivers rather than arbitrary inputs), including:
</p>
<ul>
  <li>Revenue growth trajectories (base, conservative, upside)</li>
  <li>Margin stabilization / mean reversion logic</li>
  <li>Tax rate normalization</li>
  <li>Capital intensity assumptions (CapEx and working capital behavior)</li>
</ul>
<p>
  The model is built for sensitivity testing across <b>growth</b>, <b>WACC</b>, <b>terminal assumptions</b>, and <b>multiples</b>.
</p>

<h3>5) Cost of Capital (WACC)</h3>
<p>
  The engine computes WACC using a CAPM-based cost of equity and a tax-adjusted cost of debt:
</p>
<pre><code>WACC = (E/V) * Re + (D/V) * Rd * (1 - T)</code></pre>
<ul>
  <li><b>Re:</b> derived from CAPM inputs</li>
  <li><b>Rd:</b> inferred from credit/spread logic (where applicable)</li>
  <li><b>Weights:</b> based on market capitalization + debt</li>
</ul>
<p>
  This explicitly separates <b>operating risk</b> from <b>financing structure</b>, which is critical in professional valuation.
</p>

<h3>6) Terminal Value Computation</h3>
<p>
  Two terminal value methods are supported:
</p>
<ul>
  <li><b>Gordon Growth:</b> TV = FCF<sub>n+1</sub> / (WACC - g)</li>
  <li><b>Exit Multiple:</b> TV = EBITDA<sub>n</sub> × Exit Multiple</li>
</ul>
<p>
  The engine enforces key constraints (e.g., <b>g &lt; WACC</b>) and flags unrealistic terminal dominance.
</p>

<h3>7) Enterprise Value → Equity Value</h3>
<p>
  After discounting forecast cash flows and terminal value:
</p>
<ul>
  <li><b>Enterprise Value</b> = PV(FCFs) + PV(Terminal Value)</li>
  <li><b>Equity Value</b> = Enterprise Value - Net Debt</li>
  <li><b>Implied Share Price</b> = Equity Value / Shares Outstanding</li>
</ul>
<p>
  Intermediate steps are retained for transparency and exported for auditability.
</p>

<h3>8) Excel Output Automation</h3>
<p>
  The engine exports a standardized Excel valuation output designed for readability and review. Typical sections include:
</p>
<ul>
  <li>Assumptions panel</li>
  <li>Operating projections</li>
  <li>FCFF schedule</li>
  <li>Discount factors and PV bridge</li>
  <li>Terminal value breakdown</li>
  <li>Enterprise → equity reconciliation</li>
  <li>Sensitivity tables</li>
</ul>
<p>
  The output structure mimics professional banking modeling conventions: clean segmentation, consistent labels, and traceability.
</p>

<hr />

<h2>📊 Analytical Integrity &amp; Error Handling</h2>
<p>
  The engine includes safeguards commonly needed in real valuation work:
</p>
<ul>
  <li>Division-by-zero and negative denominator protection</li>
  <li>Missing EBITDA / unstable margins fallback logic</li>
  <li>Terminal growth constraints and sanity checks</li>
  <li>Outlier detection for margin swings and cash flow instability</li>
  <li>Transparent logging of assumptions when inputs are incomplete</li>
</ul>

<hr />

<h2>🧩 Architecture (High Level)</h2>
<ul>
  <li><b>Backend:</b> Node.js valuation engine (modular financial calculation pipeline)</li>
  <li><b>Data Layer:</b> structured financial + market data retrieval (Yahoo Finance endpoints)</li>
  <li><b>Model Layer:</b> normalization → FCFF → WACC → forecasting → terminal value</li>
  <li><b>Reporting Layer:</b> standardized Excel output for auditability and repeatability</li>
</ul>
<p>
  The modular architecture supports future extensions such as comparable-company valuation, AI-assisted peer selection, and
  scenario stress testing.
</p>

<hr />

<h2>🚀 Roadmap (Planned Extensions)</h2>
<ul>
  <li>Comparable Company Valuation automation (EV/Revenue, EV/EBITDA, P/E, P/B)</li>
  <li>AI-assisted peer identification + relevance weighting</li>
  <li>Automated sensitivity heatmaps and scenario stress tests</li>
  <li>Quality filters: liquidity, profitability, size band, geography, segment similarity</li>
</ul>

<h2>⚖ Disclaimer</h2>
<p>
  This project is for educational and analytical purposes only and does not constitute investment advice.
</p>


## 👤 About the Author

<table>
  <tr>
    <td>
      <img src="https://avatars.githubusercontent.com/u/00000000?v=4" width="100" alt="Thies Boese Avatar" style="border-radius:50%;"/>
    </td>
    <td>
      <b>Thies Boese</b><br>
      <span>🎓 Finance Master's Applicant & International Business Bachelor</span><br>
      <span>🏫 <b>Maastricht University</b></span><br>
      <a href="mailto:thiesboese05@gmail.com">
        <img src="https://img.shields.io/badge/Email-D14836?style=flat&logo=gmail&logoColor=white" alt="Email"/>
      </a>
      <a href="https://www.linkedin.com/in/thies-boese">
        <img src="https://img.shields.io/badge/LinkedIn-0A66C2?style=flat&logo=linkedin&logoColor=white" alt="LinkedIn"/>
      </a>
    </td>
  </tr>
</table>

> Developed and maintained by Thies Boese as part of my academic and professional portfolio.<br>
> For networking, collaboration, or inquiries, please connect via the links above.

---
