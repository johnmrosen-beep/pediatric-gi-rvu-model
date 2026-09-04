import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Cell,
} from "recharts";

// ---------------------------------------------------------------------------
// Reference data: 2026 CMS Physician Fee Schedule, time-based office/outpatient E/M
// ---------------------------------------------------------------------------
const NEW_CODES = [
  { level: 2, code: "99202", wrvu: 0.93, minTime: 15, maxTime: 29, mdm: "Straightforward" },
  { level: 3, code: "99203", wrvu: 1.6, minTime: 30, maxTime: 44, mdm: "Low" },
  { level: 4, code: "99204", wrvu: 2.6, minTime: 45, maxTime: 59, mdm: "Moderate" },
  { level: 5, code: "99205", wrvu: 3.5, minTime: 60, maxTime: 74, mdm: "High" },
];

const EST_CODES = [
  { level: 2, code: "99212", wrvu: 0.7, minTime: 10, maxTime: 19, mdm: "Straightforward" },
  { level: 3, code: "99213", wrvu: 1.3, minTime: 20, maxTime: 29, mdm: "Low" },
  { level: 4, code: "99214", wrvu: 1.92, minTime: 30, maxTime: 39, mdm: "Moderate" },
  { level: 5, code: "99215", wrvu: 2.8, minTime: 40, maxTime: 54, mdm: "High" },
];

const levelKeys = [2, 3, 4, 5];

function defaultMix(a, b, c, d) {
  return { 2: a, 3: b, 4: c, 5: d };
}

// ---------------------------------------------------------------------------
// Default state factory
// ---------------------------------------------------------------------------
function defaultClinic(kind) {
  if (kind === "general") {
    return {
      label: "General GI",
      slotMinutes: 30,
      newShare: 0.5,
      alwaysFollowUp: false,
      newMix: defaultMix(0, 0, 0, 0),
      followUpMix: defaultMix(0, 0, 0, 0),
      newNoShowRate: 0,
      followUpNoShowRate: 0,
    };
  }
  return {
    label: "Motility",
    slotMinutes: 60,
    newShare: 0,
    alwaysFollowUp: true,
    newMix: defaultMix(0, 0, 0, 0),
    followUpMix: defaultMix(0, 0, 0, 0),
    newNoShowRate: 0,
    followUpNoShowRate: 0,
  };
}

function defaultState() {
  return {
    sessionsPerWeek: 5,
    hoursPerSession: 4,
    weeksPerYear: 47,
    baselineGeneralShare: 0.6,
    conversionFactor: 33.4009,
    facilityWrvuPlaceholder: 0,
    general: defaultClinic("general"),
    specialty: defaultClinic("specialty"),
  };
}

// ---------------------------------------------------------------------------
// Calculation helpers
// ---------------------------------------------------------------------------
function mixSum(mix) {
  return levelKeys.reduce((s, k) => s + (Number(mix[k]) || 0), 0);
}

function blendedWrvu(mix, codeTable) {
  return codeTable.reduce(
    (s, row) => s + (Number(mix[row.level]) || 0) * row.wrvu,
    0
  );
}

function atRiskShare(mix, codeTable, slotMinutes) {
  return codeTable
    .filter((row) => row.minTime > slotMinutes)
    .reduce((s, row) => s + (Number(mix[row.level]) || 0), 0);
}

function atRiskCodes(mix, codeTable, slotMinutes) {
  return codeTable.filter(
    (row) => row.minTime > slotMinutes && (Number(mix[row.level]) || 0) > 0
  );
}

function clinicMetrics(clinic, hoursPerSession) {
  const newBlended = blendedWrvu(clinic.newMix, NEW_CODES);
  const fuBlended = blendedWrvu(clinic.followUpMix, EST_CODES);
  const newShare = clinic.alwaysFollowUp ? 0 : Number(clinic.newShare) || 0;
  const overall = newShare * newBlended + (1 - newShare) * fuBlended;
  const slots = (hoursPerSession * 60) / (Number(clinic.slotMinutes) || 1);
  const noShowRate = clinic.alwaysFollowUp
    ? Number(clinic.followUpNoShowRate) || 0
    : newShare * (Number(clinic.newNoShowRate) || 0) +
      (1 - newShare) * (Number(clinic.followUpNoShowRate) || 0);
  const effectiveVisits = slots * (1 - noShowRate);
  const wrvuPerSessionTemplated = slots * overall;
  const wrvuPerSession = effectiveVisits * overall;
  const wrvuPerHour = wrvuPerSession / hoursPerSession;
  return {
    newBlended,
    fuBlended,
    overall,
    slots,
    noShowRate,
    effectiveVisits,
    wrvuPerSessionTemplated,
    wrvuPerSession,
    wrvuPerHour,
  };
}

const pctFmt = (v) => `${(v * 100).toFixed(1)}%`;
const num1 = (v) => v.toLocaleString(undefined, { maximumFractionDigits: 1 });
const num0 = (v) => v.toLocaleString(undefined, { maximumFractionDigits: 0 });
const cur0 = (v) =>
  v.toLocaleString(undefined, {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0,
  });

// ---------------------------------------------------------------------------
// Small building blocks
// ---------------------------------------------------------------------------
function Field({ label, hint, children }) {
  return (
    <label className="block">
      <span className="block text-sm text-slate-600 mb-1">{label}</span>
      {children}
      {hint && <span className="block text-xs text-slate-400 mt-1">{hint}</span>}
    </label>
  );
}

function NumberInput({ value, onChange, step = 1, min, max, suffix }) {
  return (
    <div className="flex items-center gap-2">
      <input
        type="number"
        className="w-28 border border-stone-300 rounded px-2 py-1 text-slate-900 tabular-nums focus:outline-none focus:ring-2 focus:ring-slate-400"
        value={value}
        step={step}
        min={min}
        max={max}
        onChange={(e) => onChange(parseFloat(e.target.value))}
      />
      {suffix && <span className="text-sm text-slate-500">{suffix}</span>}
    </div>
  );
}

function PercentInput({ value, onChange }) {
  return (
    <div className="flex items-center gap-1">
      <input
        type="number"
        className="w-20 border border-stone-300 rounded px-2 py-1 text-slate-900 tabular-nums focus:outline-none focus:ring-2 focus:ring-slate-400"
        value={Math.round(value * 1000) / 10}
        step={0.5}
        min={0}
        max={100}
        onChange={(e) => onChange((parseFloat(e.target.value) || 0) / 100)}
      />
      <span className="text-sm text-slate-500">%</span>
    </div>
  );
}

function SumCheck({ sum }) {
  const ok = Math.abs(sum - 1) < 0.001;
  return (
    <span
      className={
        "text-xs font-medium tabular-nums " +
        (ok ? "text-emerald-700" : "text-amber-700")
      }
    >
      {ok ? "✓ sums to 100%" : `⚠ currently ${pctFmt(sum)} — must total 100%`}
    </span>
  );
}

function MixTable({ title, mix, onChange, codeTable, accent }) {
  const sum = mixSum(mix);
  return (
    <div className="mt-4">
      <div className="flex items-baseline justify-between mb-2">
        <h4 className="text-sm font-medium text-slate-700">{title}</h4>
        <SumCheck sum={sum} />
      </div>
      <table className="w-full text-sm border-collapse">
        <thead>
          <tr className="text-left text-slate-500 border-b border-stone-200">
            <th className="py-1 font-normal">Level</th>
            <th className="py-1 font-normal">CPT</th>
            <th className="py-1 font-normal text-right">% of visits</th>
            <th className="py-1 font-normal text-right">wRVU</th>
            <th className="py-1 font-normal text-right">Weighted</th>
          </tr>
        </thead>
        <tbody>
          {codeTable.map((row) => (
            <tr key={row.code} className="border-b border-stone-100">
              <td className="py-1.5 text-slate-700 align-top">L{row.level}</td>
              <td className="py-1.5 text-slate-500 align-top">
                {row.code}
                <div className="text-xs text-slate-400 font-normal leading-snug">
                  {row.mdm} MDM
                  <br />
                  {row.minTime}–{row.maxTime} min
                </div>
              </td>
              <td className="py-1.5 text-right align-top">
                <PercentInput
                  value={mix[row.level]}
                  onChange={(v) => onChange({ ...mix, [row.level]: v })}
                />
              </td>
              <td className="py-1.5 text-right tabular-nums text-slate-500 align-top">
                {row.wrvu.toFixed(2)}
              </td>
              <td className="py-1.5 text-right tabular-nums text-slate-700 align-top">
                {(row.wrvu * (Number(mix[row.level]) || 0)).toFixed(2)}
              </td>
            </tr>
          ))}
          <tr>
            <td colSpan={4} className="py-1.5 text-right font-medium text-slate-700">
              Blended wRVU / visit
            </td>
            <td
              className={
                "py-1.5 text-right font-semibold tabular-nums " + accent
              }
            >
              {blendedWrvu(mix, codeTable).toFixed(2)}
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function ClinicPanel({ clinic, onChange, accentBorder, accentText, hoursPerSession }) {
  const metrics = clinicMetrics(clinic, hoursPerSession);
  const newRisk = atRiskCodes(clinic.newMix, NEW_CODES, clinic.slotMinutes);
  const fuRisk = atRiskCodes(clinic.followUpMix, EST_CODES, clinic.slotMinutes);
  const hasRisk = !clinic.alwaysFollowUp
    ? newRisk.length > 0 || fuRisk.length > 0
    : fuRisk.length > 0;

  return (
    <div className={"bg-white border-t-4 " + accentBorder + " border border-stone-200 rounded-sm p-5"}>
      <div className="flex items-center justify-between gap-3 mb-4">
        <input
          className={
            "font-serif text-lg font-medium bg-transparent border-b border-dashed border-stone-300 focus:outline-none focus:border-slate-400 " +
            accentText
          }
          value={clinic.label}
          onChange={(e) => onChange({ ...clinic, label: e.target.value })}
        />
        <span className="text-xs text-slate-400">editable name</span>
      </div>

      <div className="grid grid-cols-2 gap-4">
        <Field label="Templated slot length">
          <NumberInput
            value={clinic.slotMinutes}
            onChange={(v) => onChange({ ...clinic, slotMinutes: v })}
            step={5}
            min={5}
            suffix="min"
          />
        </Field>
      </div>

      <div className="mt-4 flex items-center gap-2">
        <input
          type="checkbox"
          id={`afu-${clinic.label}`}
          checked={clinic.alwaysFollowUp}
          onChange={(e) =>
            onChange({ ...clinic, alwaysFollowUp: e.target.checked })
          }
          className="h-4 w-4"
        />
        <label htmlFor={`afu-${clinic.label}`} className="text-sm text-slate-600">
          All visits are follow-up (no new-patient volume)
        </label>
      </div>

      {!clinic.alwaysFollowUp && (
        <div className="mt-3">
          <Field label="New-patient share of visits" hint="remainder is follow-up">
            <PercentInput
              value={clinic.newShare}
              onChange={(v) => onChange({ ...clinic, newShare: v })}
            />
          </Field>
        </div>
      )}

      {!clinic.alwaysFollowUp && (
        <MixTable
          title="E&M code mix — new patient (99202–99205)"
          mix={clinic.newMix}
          onChange={(m) => onChange({ ...clinic, newMix: m })}
          codeTable={NEW_CODES}
          accent={accentText}
        />
      )}

      <MixTable
        title="E&M code mix — follow-up (99212–99215)"
        mix={clinic.followUpMix}
        onChange={(m) => onChange({ ...clinic, followUpMix: m })}
        codeTable={EST_CODES}
        accent={accentText}
      />

      <div className="mt-5 pt-4 border-t border-stone-200">
        <h4 className="text-sm font-medium text-slate-700 mb-1">
          Historical no-show / DNKA rate
        </h4>
        <p className="text-xs text-slate-400 mb-3">
          Share of templated slots that don't convert to a completed,
          billable visit. Replace with actual rates from your scheduling
          system.
        </p>
        <div className={"grid gap-4 " + (clinic.alwaysFollowUp ? "grid-cols-1" : "grid-cols-2")}>
          {!clinic.alwaysFollowUp && (
            <Field label="New-patient no-show rate">
              <PercentInput
                value={clinic.newNoShowRate}
                onChange={(v) => onChange({ ...clinic, newNoShowRate: v })}
              />
            </Field>
          )}
          <Field label="Follow-up no-show rate">
            <PercentInput
              value={clinic.followUpNoShowRate}
              onChange={(v) => onChange({ ...clinic, followUpNoShowRate: v })}
            />
          </Field>
        </div>
      </div>

      {hasRisk && (
        <div className="mt-4 bg-amber-50 border border-amber-300 rounded-sm px-3 py-2 text-xs text-amber-800">
          <span className="font-medium">Slot-length flag: </span>
          {[...newRisk, ...fuRisk].map((r) => r.code).join(", ")} require(s)
          more minutes than the {clinic.slotMinutes}-minute template allows.
          Confirm total time is genuinely logged past the scheduled block, or
          revisit the template length.
        </div>
      )}

      <div className="mt-4 pt-3 border-t border-stone-200 grid grid-cols-2 gap-4 text-sm">
        <div>
          <div className="text-slate-500">Templated slots / session</div>
          <div className="tabular-nums text-slate-700">{num1(metrics.slots)}</div>
        </div>
        <div>
          <div className="text-slate-500">Effective (completed) visits / session</div>
          <div className={"tabular-nums font-medium " + accentText}>
            {num1(metrics.effectiveVisits)}
            <span className="text-slate-400 font-normal">
              {" "}
              (blended no-show {pctFmt(metrics.noShowRate)})
            </span>
          </div>
        </div>
        <div>
          <div className="text-slate-500">wRVU / session</div>
          <div className={"tabular-nums text-lg font-semibold " + accentText}>
            {metrics.wrvuPerSession.toFixed(2)}
            <span className="text-xs text-slate-400 font-normal">
              {" "}
              / {metrics.wrvuPerSessionTemplated.toFixed(2)} templated
            </span>
          </div>
        </div>
        <div>
          <div className="text-slate-500">wRVU / clinic hour</div>
          <div className={"tabular-nums text-lg font-semibold " + accentText}>
            {metrics.wrvuPerHour.toFixed(2)}
          </div>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Reference tab
// ---------------------------------------------------------------------------
function ReferenceTab() {
  const rows = [...NEW_CODES, ...EST_CODES];
  return (
    <div className="bg-white border border-stone-200 rounded-sm p-5">
      <h3 className="font-serif text-lg text-slate-800 mb-1">
        2026 CMS Physician Fee Schedule — time-based E/M reference
      </h3>
      <p className="text-sm text-slate-500 mb-4">
        Work RVUs for 99202–99215, unchanged from the 2021 office-visit
        revaluation and exempt from the CY2026 –2.5% efficiency adjustment
        (E/M codes are classified as time-based).
      </p>
      <table className="w-full text-sm border-collapse">
        <thead>
          <tr className="text-left text-slate-500 border-b border-stone-200">
            <th className="py-2 font-normal">CPT</th>
            <th className="py-2 font-normal">Descriptor</th>
            <th className="py-2 font-normal">Medical decision making</th>
            <th className="py-2 font-normal text-right">Total time (date of encounter)</th>
            <th className="py-2 font-normal text-right">Work RVU</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr
              key={r.code}
              className={
                "border-b border-stone-100 " +
                (i % 2 === 1 ? "bg-stone-50" : "")
              }
            >
              <td className="py-1.5 text-slate-700">{r.code}</td>
              <td className="py-1.5 text-slate-600">
                {i < 4 ? "New patient" : "Established patient"}, Level {r.level}
              </td>
              <td className="py-1.5 text-slate-600">{r.mdm}</td>
              <td className="py-1.5 text-right tabular-nums text-slate-600">
                {r.minTime}–{r.maxTime} min
              </td>
              <td className="py-1.5 text-right tabular-nums text-slate-800 font-medium">
                {r.wrvu.toFixed(2)}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      <p className="text-xs text-slate-400 mt-4 leading-relaxed">
        Source: CMS CY2026 Physician Fee Schedule Final Rule (CMS-1832-F),
        effective 1/1/2026. "Total time" includes non-face-to-face work on the
        date of the encounter (chart review, orders, documentation, care
        coordination). Telehealth E/M via 99202–99215 follows the same time
        thresholds (modifier 95 / POS per payer policy). Verify current values
        against the CMS PFS Look-Up Tool before use in contracts — rates can
        change annually.
      </p>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Shared scenario computation (used by the tab, Excel export, and print summary)
// ---------------------------------------------------------------------------
const SCENARIO_STEPS = [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0];

function computeModel(state) {
  const general = clinicMetrics(state.general, state.hoursPerSession);
  const specialty = clinicMetrics(state.specialty, state.hoursPerSession);
  const advantage =
    specialty.wrvuPerHour > 0 ? general.wrvuPerHour / specialty.wrvuPerHour : 0;

  const scenarioAt = (p) => {
    const genSessions = p * state.sessionsPerWeek;
    const specSessions = (1 - p) * state.sessionsPerWeek;
    const weekly =
      genSessions * general.wrvuPerSession + specSessions * specialty.wrvuPerSession;
    return { p, genSessions, specSessions, weekly, annual: weekly * state.weeksPerYear };
  };

  const scenarioRows = SCENARIO_STEPS.map(scenarioAt);
  const baseline = scenarioAt(state.baselineGeneralShare);
  const optimum = scenarioRows[scenarioRows.length - 1];
  const deltaAnnual = optimum.annual - baseline.annual;

  return { general, specialty, advantage, scenarioRows, baseline, optimum, deltaAnnual };
}

// ---------------------------------------------------------------------------
// Production model tab
// ---------------------------------------------------------------------------
function ProductionTab({ state }) {
  const { general, specialty, advantage, scenarioRows, baseline, optimum, deltaAnnual } =
    computeModel(state);

  const chartData = [
    { name: state.general.label, wrvuPerHour: general.wrvuPerHour, fill: "#7f1d1d" },
    { name: state.specialty.label, wrvuPerHour: specialty.wrvuPerHour, fill: "#115e59" },
  ];

  return (
    <div className="space-y-6">
      <div className="bg-white border border-stone-200 rounded-sm p-6">
        <div className="flex flex-wrap items-end gap-8">
          <div>
            <div className="text-sm text-slate-500">
              {state.general.label} vs. {state.specialty.label}
            </div>
            <div className="font-serif text-4xl text-slate-900 tabular-nums">
              {advantage.toFixed(2)}x
            </div>
            <div className="text-sm text-slate-500">
              wRVU-per-hour advantage, {state.general.label}
            </div>
            <div className="text-xs text-slate-400 mt-1">
              after historical no-show / DNKA rates
            </div>
          </div>
          <div className="flex-1 min-w-[260px] h-40">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartData} margin={{ top: 8, right: 8, left: 0, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#E7E5E4" vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 12, fill: "#57534E" }} axisLine={false} tickLine={false} />
                <YAxis tick={{ fontSize: 12, fill: "#57534E" }} axisLine={false} tickLine={false} width={32} />
                <Tooltip
                  formatter={(v) => [`${v.toFixed(2)} wRVU/hr`, "Production"]}
                  contentStyle={{ fontSize: 12, borderRadius: 2 }}
                />
                <Bar dataKey="wrvuPerHour" radius={[2, 2, 0, 0]}>
                  {chartData.map((d, i) => (
                    <Cell key={i} fill={d.fill} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="bg-white border border-stone-200 rounded-sm p-5">
        <h3 className="font-serif text-lg text-slate-800 mb-1">
          Weekly / annual production by schedule mix
        </h3>
        <p className="text-sm text-slate-500 mb-4">
          Share of weekly sessions run as {state.general.label} vs.{" "}
          {state.specialty.label}.
        </p>
        <div className="overflow-x-auto">
          <table className="w-full text-sm border-collapse min-w-[560px]">
            <thead>
              <tr className="text-left text-slate-500 border-b border-stone-200">
                <th className="py-2 font-normal">% {state.general.label}</th>
                <th className="py-2 font-normal text-right">{state.general.label} sessions/wk</th>
                <th className="py-2 font-normal text-right">{state.specialty.label} sessions/wk</th>
                <th className="py-2 font-normal text-right">Weekly wRVU</th>
                <th className="py-2 font-normal text-right">Annual wRVU</th>
                <th className="py-2 font-normal text-right">Annual $ (ref. only)</th>
              </tr>
            </thead>
            <tbody>
              {scenarioRows.map((row, i) => (
                <tr
                  key={row.p}
                  className={
                    "border-b border-stone-100 " + (i % 2 === 1 ? "bg-stone-50" : "")
                  }
                >
                  <td className="py-1.5 tabular-nums">{pctFmt(row.p)}</td>
                  <td className="py-1.5 text-right tabular-nums">{num1(row.genSessions)}</td>
                  <td className="py-1.5 text-right tabular-nums">{num1(row.specSessions)}</td>
                  <td className="py-1.5 text-right tabular-nums">{num1(row.weekly)}</td>
                  <td className="py-1.5 text-right tabular-nums">{num0(row.annual)}</td>
                  <td className="py-1.5 text-right tabular-nums text-slate-500">
                    {cur0(row.annual * state.conversionFactor)}
                  </td>
                </tr>
              ))}
              <tr className="border-t-2 border-stone-300 bg-amber-50">
                <td className="py-2 font-medium tabular-nums">{pctFmt(baseline.p)}</td>
                <td className="py-2 text-right tabular-nums">{num1(baseline.genSessions)}</td>
                <td className="py-2 text-right tabular-nums">{num1(baseline.specSessions)}</td>
                <td className="py-2 text-right tabular-nums">{num1(baseline.weekly)}</td>
                <td className="py-2 text-right tabular-nums font-medium">{num0(baseline.annual)}</td>
                <td className="py-2 text-right tabular-nums text-slate-500">
                  {cur0(baseline.annual * state.conversionFactor)}
                </td>
              </tr>
              <tr className="text-emerald-800 font-medium">
                <td className="py-2 tabular-nums">100.0%</td>
                <td className="py-2 text-right tabular-nums">{num1(optimum.genSessions)}</td>
                <td className="py-2 text-right tabular-nums">{num1(optimum.specSessions)}</td>
                <td className="py-2 text-right tabular-nums">{num1(optimum.weekly)}</td>
                <td className="py-2 text-right tabular-nums">{num0(optimum.annual)}</td>
                <td className="py-2 text-right tabular-nums">
                  {cur0(optimum.annual * state.conversionFactor)}
                </td>
              </tr>
            </tbody>
          </table>
        </div>
        <p className="text-sm text-slate-600 mt-4">
          Baseline (amber row) → 100% {state.general.label} (green row):{" "}
          <span className="font-medium tabular-nums">
            {deltaAnnual >= 0 ? "+" : ""}
            {num0(deltaAnnual)} wRVU/year
          </span>
          .
        </p>
      </div>

      <div className="bg-white border border-stone-200 rounded-sm p-5 text-sm text-slate-600 leading-relaxed">
        <span className="font-medium text-slate-800">Not yet modeled: </span>
        facility RVU / facility revenue for these telehealth visits. This
        depends on the clinic's provider-based billing status and
        payer-specific telehealth facility-fee policy. Confirm with revenue
        cycle / coding compliance before layering in total (direct +
        downstream) revenue.
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Assumptions tab
// ---------------------------------------------------------------------------
function AssumptionsTab({ state, setState }) {
  const totalHours = state.sessionsPerWeek * state.hoursPerSession;
  return (
    <div className="space-y-6">
      <div className="bg-white border border-stone-200 rounded-sm p-5">
        <h3 className="font-serif text-lg text-slate-800 mb-4">Clinic schedule</h3>
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          <Field label="Sessions / week">
            <NumberInput
              value={state.sessionsPerWeek}
              onChange={(v) => setState({ ...state, sessionsPerWeek: v })}
              step={1}
              min={1}
            />
          </Field>
          <Field label="Hours / session">
            <NumberInput
              value={state.hoursPerSession}
              onChange={(v) => setState({ ...state, hoursPerSession: v })}
              step={0.5}
              min={0.5}
            />
          </Field>
          <Field label="Clinic weeks / year" hint="excl. holidays, CME, PTO">
            <NumberInput
              value={state.weeksPerYear}
              onChange={(v) => setState({ ...state, weeksPerYear: v })}
              step={1}
              min={1}
              max={52}
            />
          </Field>
          <Field label="Total clinic hours / week" hint="computed">
            <div className="tabular-nums text-slate-700 py-1">{num1(totalHours)}</div>
          </Field>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
        <ClinicPanel
          clinic={state.general}
          onChange={(c) => setState({ ...state, general: c })}
          accentBorder="border-t-red-800"
          accentText="text-red-800"
          hoursPerSession={state.hoursPerSession}
        />
        <ClinicPanel
          clinic={state.specialty}
          onChange={(c) => setState({ ...state, specialty: c })}
          accentBorder="border-t-teal-800"
          accentText="text-teal-800"
          hoursPerSession={state.hoursPerSession}
        />
      </div>

      <div className="bg-white border border-stone-200 rounded-sm p-5">
        <h3 className="font-serif text-lg text-slate-800 mb-4">
          Current baseline &amp; reference values
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <Field
            label={`Current share of sessions as ${state.general.label}`}
            hint="vs. specialty — sets the amber baseline row"
          >
            <PercentInput
              value={state.baselineGeneralShare}
              onChange={(v) => setState({ ...state, baselineGeneralShare: v })}
            />
          </Field>
          <Field
            label="Medicare conversion factor"
            hint="2026 non-QP, national, pre-GPCI — reference only"
          >
            <NumberInput
              value={state.conversionFactor}
              onChange={(v) => setState({ ...state, conversionFactor: v })}
              step={0.01}
              suffix="$/RVU"
            />
          </Field>
          <Field
            label="Facility RVU / encounter"
            hint="not yet modeled — confirm billing status first"
          >
            <NumberInput
              value={state.facilityWrvuPlaceholder}
              onChange={(v) => setState({ ...state, facilityWrvuPlaceholder: v })}
              step={0.1}
              min={0}
            />
          </Field>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// State ↔ URL (hash fragment) — no backend, no server ever sees this
// ---------------------------------------------------------------------------
// The fragment ("#...") is never sent in the HTTP request for the page, so
// unlike a query string it never reaches a server log, CDN, or analytics
// pipeline — it only ever exists in the browser. That also means every
// user's data is isolated by construction: there is no shared store a
// concurrent user could read from, and the only way anyone sees your inputs
// is if you hand them this exact URL. It is not encrypted, though —
// anyone with the link can decode it, so treat the link like you would any
// other "shareable by URL" tool.
function toBase64Url(str) {
  const bytes = new TextEncoder().encode(str);
  let binary = "";
  bytes.forEach((b) => (binary += String.fromCharCode(b)));
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function fromBase64Url(b64url) {
  let b64 = b64url.replace(/-/g, "+").replace(/_/g, "/");
  while (b64.length % 4) b64 += "=";
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}

function encodeStateToHash(state) {
  try {
    return "#" + toBase64Url(JSON.stringify(state));
  } catch (e) {
    return "";
  }
}

function decodeStateFromHash() {
  const raw = (window.location.hash || "").replace(/^#/, "");
  if (!raw) return null;
  try {
    const parsed = JSON.parse(fromBase64Url(raw));
    const def = defaultState();
    // Shallow-merge onto defaults, and deep-merge each clinic, so an older
    // link (missing a field added in a later version) still loads cleanly.
    return {
      ...def,
      ...parsed,
      general: { ...def.general, ...(parsed.general || {}) },
      specialty: { ...def.specialty, ...(parsed.specialty || {}) },
    };
  } catch (e) {
    console.warn("Could not read assumptions from the link; using defaults.", e);
    return null;
  }
}

// ---------------------------------------------------------------------------
// Excel export (SheetJS) — snapshot of current inputs + computed results
// ---------------------------------------------------------------------------
function buildAssumptionsRows(state) {
  const rows = [];
  rows.push(["Pediatric GI Clinic — RVU Production Model — Assumptions Snapshot"]);
  rows.push([`Exported ${new Date().toLocaleString()}`]);
  rows.push([]);
  rows.push(["Clinic Schedule"]);
  rows.push(["Sessions per week", state.sessionsPerWeek]);
  rows.push(["Hours per session", state.hoursPerSession]);
  rows.push(["Clinic weeks per year", state.weeksPerYear]);
  rows.push(["Total clinic hours per week", state.sessionsPerWeek * state.hoursPerSession]);
  rows.push([]);

  [state.general, state.specialty].forEach((clinic) => {
    rows.push([clinic.label]);
    rows.push(["Templated slot length (min)", clinic.slotMinutes]);
    rows.push(["All visits are follow-up", clinic.alwaysFollowUp ? "Yes" : "No"]);
    if (!clinic.alwaysFollowUp) {
      rows.push(["New-patient share of visits", clinic.newShare]);
      rows.push(["New-patient no-show / DNKA rate", clinic.newNoShowRate]);
    }
    rows.push(["Follow-up no-show / DNKA rate", clinic.followUpNoShowRate]);
    rows.push([]);
    if (!clinic.alwaysFollowUp) {
      rows.push(["New-patient E&M mix (99202–99205)", "Level", "CPT", "% of visits", "wRVU"]);
      NEW_CODES.forEach((c) =>
        rows.push(["", `Level ${c.level}`, c.code, clinic.newMix[c.level], c.wrvu])
      );
      rows.push([]);
    }
    rows.push(["Follow-up E&M mix (99212–99215)", "Level", "CPT", "% of visits", "wRVU"]);
    EST_CODES.forEach((c) =>
      rows.push(["", `Level ${c.level}`, c.code, clinic.followUpMix[c.level], c.wrvu])
    );
    rows.push([]);
  });

  rows.push(["Current baseline share of sessions — " + state.general.label, state.baselineGeneralShare]);
  rows.push(["Medicare conversion factor ($/RVU, 2026 non-QP, reference only)", state.conversionFactor]);
  rows.push(["Facility RVU / encounter (not yet modeled)", state.facilityWrvuPlaceholder]);
  return rows;
}

function buildReferenceRows() {
  const rows = [["CPT Code", "Descriptor", "Total Time Threshold (date of encounter)", "Work RVU (2026, non-facility)"]];
  NEW_CODES.forEach((c) =>
    rows.push([c.code, `New patient, Level ${c.level}`, `${c.minTime}–${c.maxTime} min`, c.wrvu])
  );
  EST_CODES.forEach((c) =>
    rows.push([c.code, `Established patient, Level ${c.level}`, `${c.minTime}–${c.maxTime} min`, c.wrvu])
  );
  rows.push([]);
  rows.push([
    "Source: CMS CY2026 Physician Fee Schedule Final Rule (CMS-1832-F), effective 1/1/2026. " +
      "Work RVUs unchanged from the 2021 E/M revaluation; exempt from the CY2026 -2.5% non-time-based " +
      "efficiency adjustment. Verify against the CMS PFS Look-Up Tool before use in contracts.",
  ]);
  return rows;
}

function buildProductionRows(state) {
  const { general, specialty, advantage, scenarioRows, baseline, optimum, deltaAnnual } =
    computeModel(state);
  const rows = [];
  rows.push(["Metric", state.general.label, state.specialty.label]);
  rows.push(["Blended wRVU / visit", general.overall, specialty.overall]);
  rows.push(["Templated slots / session", general.slots, specialty.slots]);
  rows.push(["Blended no-show / DNKA rate", general.noShowRate, specialty.noShowRate]);
  rows.push(["Effective (completed) visits / session", general.effectiveVisits, specialty.effectiveVisits]);
  rows.push(["wRVU / session (templated, pre no-show)", general.wrvuPerSessionTemplated, specialty.wrvuPerSessionTemplated]);
  rows.push(["wRVU / session (effective, after no-show)", general.wrvuPerSession, specialty.wrvuPerSession]);
  rows.push(["wRVU / clinic hour (effective)", general.wrvuPerHour, specialty.wrvuPerHour]);
  rows.push([]);
  rows.push([`${state.general.label} wRVU/hr advantage over ${state.specialty.label}`, advantage]);
  rows.push([]);
  rows.push(["Weekly / Annual Production by Schedule Mix"]);
  rows.push([
    "Scenario",
    `% ${state.general.label}`,
    `${state.general.label} sessions/wk`,
    `${state.specialty.label} sessions/wk`,
    "Weekly wRVU",
    "Annual wRVU",
    "Annual $ (Medicare CF, reference only)",
  ]);
  scenarioRows.forEach((row) => {
    rows.push([
      `${(row.p * 100).toFixed(0)}%`,
      row.p,
      row.genSessions,
      row.specSessions,
      row.weekly,
      row.annual,
      row.annual * state.conversionFactor,
    ]);
  });
  rows.push([
    "Current baseline",
    baseline.p,
    baseline.genSessions,
    baseline.specSessions,
    baseline.weekly,
    baseline.annual,
    baseline.annual * state.conversionFactor,
  ]);
  rows.push([
    `100% ${state.general.label} (optimum)`,
    optimum.p,
    optimum.genSessions,
    optimum.specSessions,
    optimum.weekly,
    optimum.annual,
    optimum.annual * state.conversionFactor,
  ]);
  rows.push([]);
  rows.push(["Incremental annual wRVU: baseline → optimum", deltaAnnual]);
  rows.push([]);
  rows.push([
    "Facility RVU / facility revenue is not included. Depends on provider-based billing status and " +
      "payer-specific telehealth facility-fee policy — confirm with revenue cycle before adding.",
  ]);
  return rows;
}

function handleExportExcel(state) {
  const wb = XLSX.utils.book_new();

  const wsA = XLSX.utils.aoa_to_sheet(buildAssumptionsRows(state));
  wsA["!cols"] = [{ wch: 40 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 8 }];
  XLSX.utils.book_append_sheet(wb, wsA, "Assumptions");

  const wsR = XLSX.utils.aoa_to_sheet(buildReferenceRows());
  wsR["!cols"] = [{ wch: 10 }, { wch: 28 }, { wch: 26 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, wsR, "CPT_Reference");

  const wsP = XLSX.utils.aoa_to_sheet(buildProductionRows(state));
  wsP["!cols"] = [{ wch: 28 }, { wch: 16 }, { wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, wsP, "Production_Model");

  const stamp = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `pediatric-gi-rvu-model-${stamp}.xlsx`);
}

// ---------------------------------------------------------------------------
// Print / PDF summary — hidden on screen, shown only when printing
// (use the browser's "Save as PDF" print destination)
// ---------------------------------------------------------------------------
function PrintSummary({ state }) {
  const { general, specialty, advantage, scenarioRows, baseline, optimum, deltaAnnual } =
    computeModel(state);

  return (
    <div className="hidden print:block text-slate-900 text-sm">
      <style>{`
        @media print {
          @page { size: auto; margin: 14mm; }
        }
      `}</style>
      <h1 className="font-serif text-2xl mb-1">Pediatric GI Clinic — RVU Production Model</h1>
      <p className="text-xs text-slate-500 mb-4">
        Summary generated {new Date().toLocaleString()} — Model 1: professional wRVU production, all-telehealth, time-based billing.
      </p>

      <h2 className="font-serif text-lg mt-4 mb-2">Headline</h2>
      <p className="mb-3">
        <span className="font-semibold">{advantage.toFixed(2)}x</span> wRVU-per-hour advantage,{" "}
        {state.general.label} vs. {state.specialty.label} (after historical no-show / DNKA rates).
      </p>

      <h2 className="font-serif text-lg mt-4 mb-2">Clinic assumptions</h2>
      <table className="w-full border-collapse mb-4">
        <thead>
          <tr className="border-b border-slate-400 text-left">
            <th className="py-1 pr-2">Metric</th>
            <th className="py-1 pr-2">{state.general.label}</th>
            <th className="py-1 pr-2">{state.specialty.label}</th>
          </tr>
        </thead>
        <tbody>
          {[
            ["Templated slot length", `${state.general.slotMinutes} min`, `${state.specialty.slotMinutes} min`],
            ["Blended wRVU / visit", general.overall.toFixed(2), specialty.overall.toFixed(2)],
            ["Blended no-show / DNKA rate", pctFmt(general.noShowRate), pctFmt(specialty.noShowRate)],
            ["Effective visits / session", num1(general.effectiveVisits), num1(specialty.effectiveVisits)],
            ["wRVU / session (effective)", general.wrvuPerSession.toFixed(2), specialty.wrvuPerSession.toFixed(2)],
            ["wRVU / clinic hour (effective)", general.wrvuPerHour.toFixed(2), specialty.wrvuPerHour.toFixed(2)],
          ].map((r) => (
            <tr key={r[0]} className="border-b border-slate-200">
              <td className="py-1 pr-2 text-slate-600">{r[0]}</td>
              <td className="py-1 pr-2 tabular-nums">{r[1]}</td>
              <td className="py-1 pr-2 tabular-nums">{r[2]}</td>
            </tr>
          ))}
        </tbody>
      </table>

      <h2 className="font-serif text-lg mt-4 mb-2">Weekly / annual production by schedule mix</h2>
      <table className="w-full border-collapse text-xs mb-4">
        <thead>
          <tr className="border-b border-slate-400 text-left">
            <th className="py-1 pr-2">% {state.general.label}</th>
            <th className="py-1 pr-2 text-right">Weekly wRVU</th>
            <th className="py-1 pr-2 text-right">Annual wRVU</th>
            <th className="py-1 pr-2 text-right">Annual $ (ref.)</th>
          </tr>
        </thead>
        <tbody>
          {scenarioRows.map((row) => (
            <tr key={row.p} className="border-b border-slate-100">
              <td className="py-0.5 pr-2 tabular-nums">{pctFmt(row.p)}</td>
              <td className="py-0.5 pr-2 text-right tabular-nums">{num1(row.weekly)}</td>
              <td className="py-0.5 pr-2 text-right tabular-nums">{num0(row.annual)}</td>
              <td className="py-0.5 pr-2 text-right tabular-nums">{cur0(row.annual * state.conversionFactor)}</td>
            </tr>
          ))}
          <tr className="border-t border-slate-400 font-medium">
            <td className="py-0.5 pr-2 tabular-nums">{pctFmt(baseline.p)} (baseline)</td>
            <td className="py-0.5 pr-2 text-right tabular-nums">{num1(baseline.weekly)}</td>
            <td className="py-0.5 pr-2 text-right tabular-nums">{num0(baseline.annual)}</td>
            <td className="py-0.5 pr-2 text-right tabular-nums">{cur0(baseline.annual * state.conversionFactor)}</td>
          </tr>
          <tr className="font-medium">
            <td className="py-0.5 pr-2 tabular-nums">100% (optimum)</td>
            <td className="py-0.5 pr-2 text-right tabular-nums">{num1(optimum.weekly)}</td>
            <td className="py-0.5 pr-2 text-right tabular-nums">{num0(optimum.annual)}</td>
            <td className="py-0.5 pr-2 text-right tabular-nums">{cur0(optimum.annual * state.conversionFactor)}</td>
          </tr>
        </tbody>
      </table>
      <p className="mb-4">
        Baseline → 100% {state.general.label}:{" "}
        <span className="font-medium tabular-nums">
          {deltaAnnual >= 0 ? "+" : ""}
          {num0(deltaAnnual)} wRVU/year
        </span>
      </p>

      <p className="text-xs text-slate-500 leading-relaxed border-t border-slate-300 pt-3">
        Professional wRVU only — facility RVU and total (direct + downstream) revenue are not
        modeled pending confirmation of provider-based billing status with revenue cycle. Conversion
        factor is the 2026 CMS non-QP national rate, pre-GPCI, for reference only — not a UAMS
        compensation-plan rate.
      </p>
    </div>
  );
}



// ---------------------------------------------------------------------------
// App shell
// ---------------------------------------------------------------------------
const TABS = ["Assumptions", "Reference", "Production Model"];

export default function App() {
  const [state, setState] = useState(() => decodeStateFromHash() || defaultState());
  const [tab, setTab] = useState("Assumptions");
  const [linkCopied, setLinkCopied] = useState(false);
  const [showLinkFallback, setShowLinkFallback] = useState(false);
  const fileInputRef = useRef(null);
  const syncTimer = useRef(null);

  // Keep the address bar's hash in sync with state (debounced) so refreshing
  // the tab, bookmarking, or copying the URL bar directly always reflects
  // the current inputs — not just whatever was there at last export.
  useEffect(() => {
    if (syncTimer.current) clearTimeout(syncTimer.current);
    syncTimer.current = setTimeout(() => {
      const hash = encodeStateToHash(state);
      if (hash && hash !== window.location.hash) {
        window.history.replaceState(null, "", hash);
      }
    }, 300);
    return () => clearTimeout(syncTimer.current);
  }, [state]);

  const handleCopyLink = async () => {
    const url = window.location.href;
    try {
      await navigator.clipboard.writeText(url);
      setLinkCopied(true);
      setShowLinkFallback(false);
      setTimeout(() => setLinkCopied(false), 1800);
    } catch (e) {
      setShowLinkFallback(true);
    }
  };

  const handleExportJSON = () => {
    const blob = new Blob([JSON.stringify(state, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "pediatric-gi-rvu-assumptions.json";
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleImportFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        setState({ ...defaultState(), ...parsed });
      } catch (err) {
        alert("Could not read that file — make sure it's a JSON export from this tool.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  return (
    <div className="min-h-screen bg-stone-50 text-slate-900">
      <header className="border-b border-stone-200 bg-white print:hidden">
        <div className="max-w-5xl mx-auto px-6 py-5 flex flex-wrap items-start justify-between gap-3">
          <div>
            <h1 className="font-serif text-2xl text-slate-900">
              Clinic Production Model
            </h1>
          </div>
          <div className="flex flex-col items-end gap-1.5">
            <div className="flex gap-2">
              <button
                onClick={handleCopyLink}
                className="text-sm bg-red-800 text-white rounded-sm px-3 py-1.5 hover:bg-red-900"
              >
                {linkCopied ? "Link copied ✓" : "Copy link"}
              </button>
              <button
                onClick={() => handleExportExcel(state)}
                className="text-sm bg-slate-900 text-white rounded-sm px-3 py-1.5 hover:bg-slate-700"
              >
                Export Excel
              </button>
              <button
                onClick={() => window.print()}
                className="text-sm bg-slate-900 text-white rounded-sm px-3 py-1.5 hover:bg-slate-700"
              >
                Export PDF summary
              </button>
            </div>
            {showLinkFallback && (
              <div className="flex items-center gap-1.5 mt-1">
                <input
                  readOnly
                  value={typeof window !== "undefined" ? window.location.href : ""}
                  onFocus={(e) => e.target.select()}
                  className="text-xs border border-stone-300 rounded-sm px-2 py-1 w-64 tabular-nums"
                />
                <button
                  onClick={() => setShowLinkFallback(false)}
                  className="text-xs text-slate-400 hover:text-slate-600"
                >
                  close
                </button>
              </div>
            )}
            <div className="flex gap-3">
              <button
                onClick={handleExportJSON}
                className="text-xs text-slate-400 underline hover:text-slate-600"
              >
                Save assumptions (JSON)
              </button>
              <button
                onClick={() => fileInputRef.current && fileInputRef.current.click()}
                className="text-xs text-slate-400 underline hover:text-slate-600"
              >
                Load assumptions (JSON)
              </button>
            </div>
            <input
              ref={fileInputRef}
              type="file"
              accept="application/json"
              className="hidden"
              onChange={handleImportFile}
            />
          </div>
        </div>
        <nav className="max-w-5xl mx-auto px-6 flex gap-6">
          {TABS.map((t) => (
            <button
              key={t}
              onClick={() => setTab(t)}
              className={
                "py-3 text-sm border-b-2 -mb-px " +
                (tab === t
                  ? "border-red-800 text-slate-900 font-medium"
                  : "border-transparent text-slate-500 hover:text-slate-700")
              }
            >
              {t}
            </button>
          ))}
        </nav>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-6 print:hidden">
        {tab === "Assumptions" && <AssumptionsTab state={state} setState={setState} />}
        {tab === "Reference" && <ReferenceTab />}
        {tab === "Production Model" && <ProductionTab state={state} />}
      </main>

      <footer className="max-w-5xl mx-auto px-6 pb-8 text-xs text-slate-400 print:hidden">
        Model 1 — professional wRVU production only. All-telehealth,
        time-based billing. Facility RVU and total-revenue views are planned
        for a later model. Your assumptions live in this page's URL (not on
        any server), so bookmarking or copying the address bar preserves
        them — "Copy link" does that in one click. Nothing is shared with
        other users unless you hand them that link.
      </footer>

      <PrintSummary state={state} />
    </div>
  );
}
