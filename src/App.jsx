import React, { useState, useCallback, useRef, Component } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import {
  UploadCloud,
  FileJson,
  FileSpreadsheet,
  CircleAlert,
  TriangleAlert,
  Info,
  CircleCheck,
  RotateCcw,
  Radio,
  ChevronDown,
  Download,
  FileDown,
} from "lucide-react";

/* ---------------------------------------------------------------------- */
/*  Sample data — a deliberately imperfect catalog so the tool has        */
/*  something to say the moment someone tries the demo.                   */
/* ---------------------------------------------------------------------- */

const SAMPLE_CATALOG = {
  productSpecification: [
    { id: "SPEC-5G-CORE", name: "5G Core SIM Spec", lifecycleStatus: "Active" },
    { id: "SPEC-ROUTER-X2", name: "Router X2 Spec", lifecycleStatus: "Active" },
    { id: "SPEC-LEGACY-DSL", name: "Legacy DSL Modem Spec", lifecycleStatus: "Retired" },
    { id: "SPEC-UNUSED", name: "Unused IoT Spec", lifecycleStatus: "Active" },
  ],
  category: [
    { id: "CAT-MOBILE", name: "Mobile", parentCategoryId: "" },
    { id: "CAT-BROADBAND", name: "Broadband", parentCategoryId: "" },
    { id: "CAT-5G-PLANS", name: "5G Plans", parentCategoryId: "CAT-MOBILE" },
    { id: "CAT-ORPHAN", name: "Orphan Add-ons", parentCategoryId: "CAT-DOES-NOT-EXIST" },
    { id: "CAT-LOOP-A", name: "Loop A", parentCategoryId: "CAT-LOOP-B" },
    { id: "CAT-LOOP-B", name: "Loop B", parentCategoryId: "CAT-LOOP-A" },
  ],
  productOfferingPrice: [
    { id: "POP-499", name: "₹499/mo", lifecycleStatus: "Active" },
    { id: "POP-999", name: "₹999/mo", lifecycleStatus: "Active" },
    { id: "POP-UNUSED", name: "₹199/mo (unused)", lifecycleStatus: "Active" },
  ],
  productOffering: [
    {
      id: "PO-5G-UNLIMITED",
      name: "5G Unlimited Plan",
      lifecycleStatus: "Active",
      validFrom: "2026-01-01",
      validTo: "2027-01-01",
      specId: "SPEC-5G-CORE",
      priceId: "POP-999",
      categoryId: "CAT-5G-PLANS",
      childOfferingIds: "",
    },
    {
      id: "PO-BROADBAND-FIBER",
      name: "Fiber Broadband 100Mbps",
      lifecycleStatus: "Active",
      validFrom: "2026-02-01",
      validTo: "2026-01-01",
      specId: "SPEC-ROUTER-X2",
      priceId: "POP-499",
      categoryId: "CAT-BROADBAND",
      childOfferingIds: "",
    },
    {
      id: "PO-DSL-BASIC",
      name: "DSL Basic",
      lifecycleStatus: "Active",
      validFrom: "2025-01-01",
      validTo: "2026-12-31",
      specId: "SPEC-LEGACY-DSL",
      priceId: "POP-499",
      categoryId: "CAT-BROADBAND",
      childOfferingIds: "",
    },
    {
      id: "PO-BUNDLE-HOME",
      name: "Home Bundle (5G + Fiber)",
      lifecycleStatus: "Active",
      validFrom: "2026-01-01",
      validTo: "2027-01-01",
      specId: "SPEC-5G-CORE",
      priceId: "POP-999",
      categoryId: "CAT-5G-PLANS",
      childOfferingIds: "PO-5G-UNLIMITED,PO-BROADBAND-FIBER,PO-MISSING-CHILD",
    },
    {
      id: "PO-NO-PRICE",
      name: "Trial SIM Offer",
      lifecycleStatus: "Active",
      validFrom: "2026-01-01",
      validTo: "",
      specId: "SPEC-5G-CORE",
      priceId: "",
      categoryId: "CAT-MOBILE",
      childOfferingIds: "",
    },
    {
      id: "PO-DUP-ID",
      name: "Duplicate ID Demo",
      lifecycleStatus: "Active",
      validFrom: "2026-01-01",
      validTo: "2027-01-01",
      specId: "",
      priceId: "",
      categoryId: "",
      childOfferingIds: "",
    },
    {
      id: "PO-DUP-ID",
      name: "Duplicate ID Demo (again)",
      lifecycleStatus: "Active",
      validFrom: "2026-01-01",
      validTo: "2027-01-01",
      specId: "",
      priceId: "",
      categoryId: "",
      childOfferingIds: "",
    },
    {
      id: "PO-SELF-BUNDLE",
      name: "Self-Referencing Bundle",
      lifecycleStatus: "Active",
      validFrom: "2026-01-01",
      validTo: "2027-01-01",
      specId: "",
      priceId: "",
      categoryId: "",
      childOfferingIds: "PO-SELF-BUNDLE",
    },
    {
      id: "PO-BAD-STATUS",
      name: "Weird Status Offer",
      lifecycleStatus: "Pending Review",
      validFrom: "not-a-date",
      validTo: "2027-01-01",
      specId: "",
      priceId: "",
      categoryId: "",
      childOfferingIds: "",
    },
    {
      id: "PO-FUTURE-ACTIVE",
      name: "Future Launch Marked Active",
      lifecycleStatus: "Active",
      validFrom: "2099-01-01",
      validTo: "2100-01-01",
      specId: "",
      priceId: "",
      categoryId: "",
      childOfferingIds: "",
    },
  ],
};

/* ---------------------------------------------------------------------- */
/*  Parsing: accepts TMF620-nested JSON, flat JSON array, or CSV/XLSX      */
/* ---------------------------------------------------------------------- */

const ENTITY_TYPES = [
  { key: "productOffering", label: "Product Offering" },
  { key: "productSpecification", label: "Product Specification" },
  { key: "category", label: "Category" },
  { key: "productOfferingPrice", label: "Offering Price" },
];

// Known TMF620 lifecycle status values (case-insensitive match).
const KNOWN_LIFECYCLE_STATUSES = new Set([
  "in study",
  "in design",
  "in test",
  "active",
  "rejected",
  "launched",
  "retired",
  "obsolete",
]);

function normalizeFromBuckets(buckets) {
  const entities = [];
  if (!buckets || typeof buckets !== "object") return entities;
  ENTITY_TYPES.forEach(({ key }) => {
    const bucket = buckets[key];
    const rows = Array.isArray(bucket) ? bucket : bucket ? [bucket] : [];
    rows.forEach((raw) => {
      if (raw && typeof raw === "object") entities.push(normalizeEntity(key, raw));
    });
  });
  return entities;
}

// Maps loose/plural/alternate key spellings to the canonical bucket keys,
// e.g. a real TMF620 export might use "productOfferings" or "specs".
const BUCKET_ALIASES = {
  productoffering: "productOffering",
  productofferings: "productOffering",
  offerings: "productOffering",
  productspecification: "productSpecification",
  productspecifications: "productSpecification",
  specifications: "productSpecification",
  specs: "productSpecification",
  category: "category",
  categories: "category",
  productofferingprice: "productOfferingPrice",
  productofferingprices: "productOfferingPrice",
  prices: "productOfferingPrice",
  offeringprices: "productOfferingPrice",
};

function canonicalizeBucketKeys(data) {
  const out = {};
  Object.keys(data).forEach((k) => {
    const canon = BUCKET_ALIASES[k.toLowerCase()];
    if (canon) out[canon] = data[k];
  });
  return out;
}

function looksLikeBareEntity(data) {
  const hasBucketKey = Object.keys(data).some((k) => BUCKET_ALIASES[k.toLowerCase()]);
  return !hasBucketKey && ("id" in data || "name" in data);
}

function normalizeEntity(type, raw) {
  const specRef =
    raw.specId ||
    (raw.productSpecification && raw.productSpecification.id) ||
    "";
  const priceRef =
    raw.priceId ||
    (Array.isArray(raw.productOfferingPrice)
      ? raw.productOfferingPrice.map((p) => p.id).join(",")
      : raw.productOfferingPrice && raw.productOfferingPrice.id) ||
    "";
  const categoryRef =
    raw.categoryId ||
    (Array.isArray(raw.category)
      ? raw.category.map((c) => c.id).join(",")
      : raw.category && raw.category.id) ||
    "";
  const childOfferings =
    raw.childOfferingIds ||
    (Array.isArray(raw.bundledProductOffering)
      ? raw.bundledProductOffering.map((c) => c.id).join(",")
      : "") ||
    "";
  const validFrom =
    raw.validFrom || (raw.validFor && raw.validFor.startDateTime) || "";
  const validTo =
    raw.validTo || (raw.validFor && raw.validFor.endDateTime) || "";

  return {
    type,
    id: String(raw.id ?? "").trim(),
    name: String(raw.name ?? "").trim(),
    lifecycleStatus: String(raw.lifecycleStatus ?? "").trim(),
    validFrom: String(validFrom ?? "").trim(),
    validTo: String(validTo ?? "").trim(),
    specId: String(specRef ?? "").trim(),
    priceId: String(priceRef ?? "").trim(),
    categoryId: String(categoryRef ?? "").trim(),
    parentCategoryId: String(raw.parentCategoryId ?? "").trim(),
    childOfferingIds: String(childOfferings ?? "").trim(),
  };
}

function parseJSON(text) {
  const data = JSON.parse(text);
  if (Array.isArray(data)) {
    const buckets = {};
    data.forEach((row) => {
      if (!row || typeof row !== "object") return;
      const t = (row.type || row["@type"] || "").toString();
      const key = ENTITY_TYPES.find(
        (e) => e.key.toLowerCase() === t.toLowerCase().replace(/\s/g, "")
      );
      const bucketKey = key ? key.key : "productOffering";
      buckets[bucketKey] = buckets[bucketKey] || [];
      buckets[bucketKey].push(row);
    });
    return normalizeFromBuckets(buckets);
  }
  if (data && typeof data === "object") {
    if (looksLikeBareEntity(data)) {
      return normalizeFromBuckets({ productOffering: [data] });
    }
    return normalizeFromBuckets(canonicalizeBucketKeys(data));
  }
  return [];
}

function parseTabular(rows) {
  const buckets = {};
  rows.forEach((row) => {
    const rawType = (row.type || row.entityType || "").toString().trim();
    const match = ENTITY_TYPES.find(
      (e) =>
        e.key.toLowerCase() === rawType.toLowerCase().replace(/\s/g, "") ||
        e.label.toLowerCase() === rawType.toLowerCase()
    );
    const bucketKey = match ? match.key : "productOffering";
    buckets[bucketKey] = buckets[bucketKey] || [];
    buckets[bucketKey].push(row);
  });
  return normalizeFromBuckets(buckets);
}

async function parseFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".json")) {
    const text = await file.text();
    return parseJSON(text);
  }
  if (name.endsWith(".csv")) {
    const text = await file.text();
    const result = Papa.parse(text, { header: true, skipEmptyLines: true });
    return parseTabular(result.data);
  }
  if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    return parseTabular(rows);
  }
  throw new Error("Unsupported file type. Upload .json, .csv, or .xlsx");
}

/* ---------------------------------------------------------------------- */
/*  Validation rules                                                       */
/* ---------------------------------------------------------------------- */

function findCycleFrom(startId, nextIdsFn, byId) {
  // DFS with a path stack; returns true if a cycle is reachable from startId.
  const visited = new Set();
  const stack = [startId];
  let steps = 0;
  while (stack.length) {
    const current = stack.pop();
    steps += 1;
    if (steps > 500) return false; // safety valve against pathological input
    const nexts = nextIdsFn(current);
    for (const n of nexts) {
      if (n === startId) return true;
      if (visited.has(n)) continue;
      visited.add(n);
      if (byId.has(n)) stack.push(n);
    }
  }
  return false;
}

function runValidation(entities) {
  const issues = [];
  let checkCount = 0;

  const byId = new Map(entities.map((e) => [e.id, e]));
  const specs = entities.filter((e) => e.type === "productSpecification");
  const categories = entities.filter((e) => e.type === "category");
  const prices = entities.filter((e) => e.type === "productOfferingPrice");
  const offerings = entities.filter((e) => e.type === "productOffering");

  const flag = (severity, entity, rule, message) => {
    issues.push({
      severity,
      entityId: entity.id,
      entityType: entity.type,
      entityName: entity.name,
      rule,
      message,
    });
  };

  // --- Duplicate ID detection (across all entity types) ------------------
  const idCounts = new Map();
  entities.forEach((e) => {
    if (!e.id) return;
    idCounts.set(e.id, (idCounts.get(e.id) || 0) + 1);
  });

  entities.forEach((e) => {
    checkCount += 4;
    if (!e.id) flag("critical", e, "Missing ID", "Entity has no identifier.");
    if (!e.name)
      flag("warning", e, "Missing name", "Entity has no display name.");
    if (!e.lifecycleStatus) {
      flag(
        "warning",
        e,
        "Missing lifecycle status",
        "No lifecycleStatus set — required by TMF620."
      );
    } else if (!KNOWN_LIFECYCLE_STATUSES.has(e.lifecycleStatus.toLowerCase())) {
      flag(
        "warning",
        e,
        "Unrecognized lifecycle status",
        `lifecycleStatus "${e.lifecycleStatus}" is not a standard TMF620 value (In Study, In Design, In Test, Active, Rejected, Launched, Retired, Obsolete).`
      );
    }
    if (e.id && idCounts.get(e.id) > 1) {
      flag(
        "critical",
        e,
        "Duplicate ID",
        `ID "${e.id}" is used by ${idCounts.get(e.id)} entities. IDs must be unique.`
      );
    }
  });

  // --- Category hierarchy integrity --------------------------------------
  categories.forEach((c) => {
    checkCount += 2;
    if (c.parentCategoryId && !byId.has(c.parentCategoryId)) {
      flag(
        "critical",
        c,
        "Orphaned category",
        `parentCategoryId "${c.parentCategoryId}" does not resolve to any category.`
      );
    }
    if (c.parentCategoryId) {
      const hasCycle = findCycleFrom(
        c.id,
        (cid) => {
          const cat = byId.get(cid);
          return cat && cat.type === "category" && cat.parentCategoryId
            ? [cat.parentCategoryId]
            : [];
        },
        byId
      );
      if (hasCycle) {
        flag(
          "critical",
          c,
          "Category cycle detected",
          `Following parentCategoryId from "${c.id}" loops back to itself.`
        );
      }
    }
  });

  // --- Cross-reference tracking for orphan/unused detection --------------
  const referencedSpecIds = new Set();
  const referencedPriceIds = new Set();
  const referencedCategoryIds = new Set();
  const referencedAsParent = new Set(
    categories.filter((c) => c.parentCategoryId).map((c) => c.parentCategoryId)
  );
  const nameCounts = new Map();

  // --- Offering-level checks ----------------------------------------------
  offerings.forEach((o) => {
    checkCount += 9;

    if (o.name) {
      const key = o.name.trim().toLowerCase();
      nameCounts.set(key, (nameCounts.get(key) || 0) + 1);
    }

    if (o.specId && !byId.has(o.specId)) {
      flag(
        "critical",
        o,
        "Broken spec reference",
        `References Product Specification "${o.specId}" which does not exist.`
      );
    } else if (o.specId) {
      referencedSpecIds.add(o.specId);
      const spec = byId.get(o.specId);
      if (
        spec.lifecycleStatus.toLowerCase() === "retired" &&
        o.lifecycleStatus.toLowerCase() === "active"
      ) {
        flag(
          "critical",
          o,
          "Active offering on retired spec",
          `Offering is Active but its spec "${o.specId}" (${spec.name}) is Retired.`
        );
      }
    } else {
      flag(
        "warning",
        o,
        "No spec reference",
        "Offering has no linked Product Specification."
      );
    }

    if (!o.priceId) {
      flag(
        "warning",
        o,
        "No price reference",
        "Offering has no linked Product Offering Price."
      );
    } else {
      o.priceId.split(",").filter(Boolean).forEach((pid) => {
        const trimmed = pid.trim();
        if (!byId.has(trimmed)) {
          flag(
            "critical",
            o,
            "Broken price reference",
            `References price "${trimmed}" which does not exist.`
          );
        } else {
          referencedPriceIds.add(trimmed);
        }
      });
    }

    if (o.categoryId) {
      o.categoryId.split(",").filter(Boolean).forEach((cid) => {
        const trimmed = cid.trim();
        if (!byId.has(trimmed)) {
          flag(
            "warning",
            o,
            "Broken category reference",
            `References category "${trimmed}" which does not exist.`
          );
        } else {
          referencedCategoryIds.add(trimmed);
        }
      });
    } else {
      flag(
        "info",
        o,
        "Uncategorized offering",
        "Offering has no categoryId set."
      );
    }

    const startDate = o.validFrom ? new Date(o.validFrom) : null;
    const endDate = o.validTo ? new Date(o.validTo) : null;
    const startInvalid = o.validFrom && isNaN(startDate);
    const endInvalid = o.validTo && isNaN(endDate);

    if (startInvalid) {
      flag(
        "warning",
        o,
        "Invalid validFrom format",
        `"${o.validFrom}" could not be parsed as a date.`
      );
    }
    if (endInvalid) {
      flag(
        "warning",
        o,
        "Invalid validTo format",
        `"${o.validTo}" could not be parsed as a date.`
      );
    }

    if (!startInvalid && !endInvalid && startDate && endDate) {
      if (startDate > endDate) {
        flag(
          "critical",
          o,
          "Invalid validFor window",
          `validFrom (${o.validFrom}) is after validTo (${o.validTo}).`
        );
      }
      if (endDate < new Date() && o.lifecycleStatus.toLowerCase() === "active") {
        flag(
          "warning",
          o,
          "Expired but marked Active",
          `validTo (${o.validTo}) has passed but lifecycleStatus is still Active.`
        );
      }
    }

    if (
      !startInvalid &&
      startDate &&
      startDate > new Date() &&
      o.lifecycleStatus.toLowerCase() === "active"
    ) {
      flag(
        "warning",
        o,
        "Not yet effective but marked Active",
        `validFrom (${o.validFrom}) is in the future but lifecycleStatus is already Active.`
      );
    }

    if (o.childOfferingIds) {
      const childIds = o.childOfferingIds.split(",").filter(Boolean);
      childIds.forEach((cid) => {
        const trimmed = cid.trim();
        const child = byId.get(trimmed);
        if (!child) {
          flag(
            "critical",
            o,
            "Broken bundle child",
            `Bundle references child offering "${trimmed}" which does not exist.`
          );
        } else if (child.lifecycleStatus.toLowerCase() !== "active") {
          flag(
            "warning",
            o,
            "Inactive bundle child",
            `Bundle child "${trimmed}" (${child.name}) has status "${child.lifecycleStatus}".`
          );
        }
      });

      const hasCycle = findCycleFrom(
        o.id,
        (oid) => {
          const off = byId.get(oid);
          return off && off.type === "productOffering" && off.childOfferingIds
            ? off.childOfferingIds.split(",").map((s) => s.trim()).filter(Boolean)
            : [];
        },
        byId
      );
      if (hasCycle) {
        flag(
          "critical",
          o,
          "Bundle cycle detected",
          `Following childOfferingIds from "${o.id}" loops back to itself.`
        );
      }
    }
  });

  // --- Duplicate offering names --------------------------------------------
  offerings.forEach((o) => {
    if (!o.name) return;
    const key = o.name.trim().toLowerCase();
    if (nameCounts.get(key) > 1) {
      flag(
        "info",
        o,
        "Duplicate offering name",
        `"${o.name}" is used by ${nameCounts.get(key)} offerings.`
      );
    }
  });

  // --- Unused specs / prices / categories (informational) -----------------
  specs.forEach((s) => {
    if (!referencedSpecIds.has(s.id)) {
      flag(
        "info",
        s,
        "Unused specification",
        "No product offering references this specification."
      );
    }
  });
  prices.forEach((p) => {
    if (!referencedPriceIds.has(p.id)) {
      flag(
        "info",
        p,
        "Unused price",
        "No product offering references this price."
      );
    }
  });
  categories.forEach((c) => {
    if (!referencedCategoryIds.has(c.id) && !referencedAsParent.has(c.id)) {
      flag(
        "info",
        c,
        "Unused category",
        "No offering or subcategory references this category."
      );
    }
  });

  const scoringIssues = issues.filter((i) => i.severity !== "info");
  const score = Math.max(
    0,
    Math.round(100 - (scoringIssues.length / Math.max(checkCount, 1)) * 100 * 4)
  );

  return {
    issues,
    score: Math.min(100, score),
    summary: {
      critical: issues.filter((i) => i.severity === "critical").length,
      warning: issues.filter((i) => i.severity === "warning").length,
      info: issues.filter((i) => i.severity === "info").length,
      entitiesScanned: entities.length,
      offerings: offerings.length,
      specs: specs.length,
      categories: categories.length,
      prices: prices.length,
    },
  };
}

/* ---------------------------------------------------------------------- */
/*  Export helpers                                                         */
/* ---------------------------------------------------------------------- */

function exportCSV(result, fileName) {
  const rows = result.issues.map((i) => ({
    Severity: i.severity,
    "Entity Type": i.entityType,
    "Entity ID": i.entityId,
    "Entity Name": i.entityName,
    Rule: i.rule,
    Message: i.message,
  }));
  const csv = Papa.unparse(rows);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `compliance-report-${(fileName || "catalog").replace(/\.[^.]+$/, "")}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function exportPDF(result, fileName) {
  const doc = new jsPDF();
  const marginX = 14;
  let y = 18;

  doc.setFontSize(16);
  doc.text("Catalog Compliance Report", marginX, y);
  y += 8;
  doc.setFontSize(10);
  doc.setTextColor(100);
  doc.text(`Source: ${fileName || "catalog"}`, marginX, y);
  y += 5;
  doc.text(`Generated: ${new Date().toLocaleString()}`, marginX, y);
  y += 5;
  doc.text(
    `Score: ${result.score}/100  |  Critical: ${result.summary.critical}  |  Warnings: ${result.summary.warning}  |  Info: ${result.summary.info}`,
    marginX,
    y
  );
  y += 8;

  autoTable(doc, {
    startY: y,
    head: [["Severity", "Type", "Entity ID", "Rule", "Message"]],
    body: result.issues.map((i) => [
      i.severity,
      i.entityType,
      i.entityId || "(no id)",
      i.rule,
      i.message,
    ]),
    styles: { fontSize: 8, cellWidth: "wrap" },
    columnStyles: { 4: { cellWidth: 70 } },
    headStyles: { fillColor: [15, 23, 42] },
    didParseCell: (data) => {
      if (data.section === "body" && data.column.index === 0) {
        const val = data.cell.raw;
        if (val === "critical") data.cell.styles.textColor = [220, 38, 38];
        else if (val === "warning") data.cell.styles.textColor = [217, 119, 6];
        else data.cell.styles.textColor = [59, 130, 246];
      }
    },
  });

  doc.save(`compliance-report-${(fileName || "catalog").replace(/\.[^.]+$/, "")}.pdf`);
}

/* ---------------------------------------------------------------------- */
/*  UI pieces                                                              */
/* ---------------------------------------------------------------------- */

function Gauge({ score }) {
  const radius = 74;
  const stroke = 14;
  const circumference = Math.PI * radius; // half circle
  const pct = Math.max(0, Math.min(100, score));
  const offset = circumference - (pct / 100) * circumference;
  const color = pct >= 85 ? "#34D399" : pct >= 60 ? "#FB923C" : "#F87171";

  return (
    <div className="relative flex flex-col items-center">
      <svg width="200" height="120" viewBox="0 0 200 120">
        <path
          d="M 20 106 A 74 74 0 0 1 180 106"
          fill="none"
          stroke="#1F2937"
          strokeWidth={stroke}
          strokeLinecap="round"
        />
        <path
          d="M 20 106 A 74 74 0 0 1 180 106"
          fill="none"
          stroke={color}
          strokeWidth={stroke}
          strokeLinecap="round"
          strokeDasharray={circumference}
          strokeDashoffset={offset}
          style={{ transition: "stroke-dashoffset 0.8s ease, stroke 0.8s ease" }}
        />
      </svg>
      <div className="absolute top-[52px] flex flex-col items-center">
        <span
          className="font-[Space_Grotesk] text-4xl font-semibold tabular-nums"
          style={{ color }}
        >
          {pct}
        </span>
        <span className="text-[10px] tracking-[0.2em] text-slate-400 uppercase">
          compliance
        </span>
      </div>
    </div>
  );
}

function StatCard({ label, value, accent }) {
  return (
    <div className="flex-1 min-w-[100px] rounded-lg border border-slate-700/60 bg-[#101827] px-4 py-3">
      <div
        className="font-[Space_Grotesk] text-2xl font-semibold tabular-nums"
        style={{ color: accent || "#E7ECF3" }}
      >
        {value}
      </div>
      <div className="text-[11px] uppercase tracking-wider text-slate-400 mt-1">
        {label}
      </div>
    </div>
  );
}

const severityConfig = {
  critical: { color: "#F87171", Icon: CircleAlert, label: "Critical" },
  warning: { color: "#FB923C", Icon: TriangleAlert, label: "Warning" },
  info: { color: "#60A5FA", Icon: Info, label: "Info" },
};

function ReportRail({ issues }) {
  const [openTypes, setOpenTypes] = useState(() => new Set(ENTITY_TYPES.map((t) => t.key)));

  const toggle = (key) => {
    setOpenTypes((prev) => {
      const next = new Set(prev);
      next.has(key) ? next.delete(key) : next.add(key);
      return next;
    });
  };

  if (issues.length === 0) {
    return (
      <div className="flex items-center gap-3 rounded-lg border border-emerald-800/50 bg-emerald-950/30 px-5 py-4 mt-6">
        <CircleCheck size={20} color="#34D399" />
        <span className="text-sm text-emerald-300">
          No issues found. Every entity passed all checks.
        </span>
      </div>
    );
  }

  const grouped = ENTITY_TYPES.map((t) => ({
    ...t,
    issues: issues.filter((i) => i.entityType === t.key),
  })).filter((g) => g.issues.length > 0);

  return (
    <div className="mt-8 space-y-6">
      {grouped.map((group) => (
        <div key={group.key}>
          <button
            onClick={() => toggle(group.key)}
            className="flex items-center gap-2 text-sm font-medium text-slate-300 mb-3 hover:text-slate-100 transition-colors"
          >
            <ChevronDown
              size={16}
              className="transition-transform"
              style={{
                transform: openTypes.has(group.key) ? "rotate(0deg)" : "rotate(-90deg)",
              }}
            />
            {group.label}
            <span className="text-slate-500 font-normal">({group.issues.length})</span>
          </button>

          {openTypes.has(group.key) && (
            <div className="relative pl-6">
              <div className="absolute left-[7px] top-2 bottom-2 w-px bg-slate-700" />
              <div className="space-y-3">
                {group.issues.map((issue, idx) => {
                  const cfg = severityConfig[issue.severity];
                  return (
                    <div key={idx} className="relative">
                      <span
                        className="absolute -left-[22px] top-3 h-3 w-3 rounded-full border-2"
                        style={{ borderColor: cfg.color, backgroundColor: "#0A0F1C" }}
                      />
                      <div className="rounded-lg border border-slate-700/60 bg-[#101827] px-4 py-3">
                        <div className="flex items-center gap-2 mb-1">
                          <cfg.Icon size={14} color={cfg.color} />
                          <span
                            className="text-[11px] font-semibold uppercase tracking-wider"
                            style={{ color: cfg.color }}
                          >
                            {cfg.label}
                          </span>
                          <span className="text-slate-500 text-xs">·</span>
                          <span className="text-sm text-slate-200 font-medium">
                            {issue.rule}
                          </span>
                        </div>
                        <p className="text-sm text-slate-400 leading-snug">{issue.message}</p>
                        <p className="mt-1 font-[IBM_Plex_Mono] text-[11px] text-slate-500">
                          {issue.entityId || "(no id)"}{" "}
                          {issue.entityName && `— ${issue.entityName}`}
                        </p>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/*  Error boundary — a render crash should never produce a blank screen    */
/* ---------------------------------------------------------------------- */

class ErrorBoundary extends Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }
  static getDerivedStateFromError(error) {
    return { error };
  }
  render() {
    if (this.state.error) {
      return (
        <div className="min-h-screen w-full bg-[#0A0F1C] text-[#E7ECF3] flex items-center justify-center px-6">
          <div className="max-w-lg w-full rounded-xl border border-red-800/50 bg-red-950/20 px-6 py-6">
            <div className="flex items-center gap-2 mb-3">
              <CircleAlert size={18} color="#F87171" />
              <span className="text-sm font-semibold text-red-300">
                Something crashed while rendering the report
              </span>
            </div>
            <p className="text-xs text-slate-400 mb-3 leading-relaxed">
              This is a bug in the checker, not your file. Details below.
            </p>
            <pre className="text-[11px] text-slate-400 bg-[#0D1420] rounded-lg p-3 overflow-auto max-h-40 whitespace-pre-wrap">
              {this.state.error?.message || String(this.state.error)}
            </pre>
            <button
              onClick={() => this.setState({ error: null })}
              className="mt-4 flex items-center gap-1.5 text-xs text-[#F2A93B] hover:text-[#f7bd63] transition-colors"
            >
              <RotateCcw size={13} />
              Try again
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

/* ---------------------------------------------------------------------- */
/*  Main app                                                               */
/* ---------------------------------------------------------------------- */

function CatalogComplianceCheckerInner() {
  const [result, setResult] = useState(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef(null);

  const handleFile = useCallback(async (file) => {
    setError("");
    try {
      const entities = await parseFile(file);
      if (entities.length === 0) {
        setError("No recognizable entities found in this file.");
        return;
      }
      setResult(runValidation(entities));
      setFileName(file.name);
    } catch (e) {
      setError(e.message || "Could not parse this file.");
    }
  }, []);

  const loadSample = () => {
    setError("");
    const entities = normalizeFromBuckets(SAMPLE_CATALOG);
    setResult(runValidation(entities));
    setFileName("sample-catalog.json (demo data)");
  };

  const reset = () => {
    setResult(null);
    setFileName("");
    setError("");
  };

  const onDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    if (e.dataTransfer.files?.[0]) handleFile(e.dataTransfer.files[0]);
  };

  return (
    <div
      className="min-h-screen w-full bg-[#0A0F1C] text-[#E7ECF3]"
      style={{ fontFamily: "Inter, system-ui, sans-serif" }}
    >
      <div className="max-w-3xl mx-auto px-6 py-14">
        {/* Header */}
        <div className="flex items-center gap-2 mb-2">
          <Radio size={18} color="#F2A93B" />
          <span className="text-[11px] uppercase tracking-[0.25em] text-[#F2A93B] font-medium">
            TMF620 Audit
          </span>
        </div>
        <h1
          className="text-3xl sm:text-4xl font-semibold mb-2"
          style={{ fontFamily: "'Space Grotesk', sans-serif" }}
        >
          Catalog Compliance Checker
        </h1>
        <p className="text-slate-400 text-sm max-w-xl mb-10">
          Upload a product catalog export — TMF620 JSON or a flat CSV/Excel
          sheet — and get a severity-ranked compliance report before it hits
          production order management.
        </p>

        {!result && (
          <>
            <div
              onDragOver={(e) => {
                e.preventDefault();
                setDragging(true);
              }}
              onDragLeave={() => setDragging(false)}
              onDrop={onDrop}
              onClick={() => inputRef.current?.click()}
              className="relative cursor-pointer rounded-xl border-2 border-dashed px-8 py-14 flex flex-col items-center text-center transition-colors"
              style={{
                borderColor: dragging ? "#F2A93B" : "#243044",
                backgroundColor: dragging ? "#12192a" : "#0D1420",
              }}
            >
              <UploadCloud size={30} color="#F2A93B" className="mb-4" />
              <p className="text-sm text-slate-300 mb-1">
                Drop a catalog file here, or click to browse
              </p>
              <p className="text-xs text-slate-500">Accepts .json · .csv · .xlsx</p>
              <input
                ref={inputRef}
                type="file"
                accept=".json,.csv,.xlsx,.xls"
                className="hidden"
                onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])}
              />
            </div>

            {error && (
              <div className="mt-4 flex items-center gap-2 text-sm text-red-400">
                <CircleAlert size={16} />
                {error}
              </div>
            )}

            <div className="mt-5 flex items-center gap-3 text-sm">
              <span className="text-slate-500">No file handy?</span>
              <button
                onClick={loadSample}
                className="text-[#F2A93B] hover:text-[#f7bd63] font-medium transition-colors"
              >
                Try the sample catalog →
              </button>
            </div>

            <div className="mt-10 grid grid-cols-2 gap-4 text-xs text-slate-500">
              <div className="flex items-start gap-2">
                <FileJson size={14} className="mt-0.5 shrink-0" />
                <span>
                  JSON: array or object of productOffering / productSpecification
                  / category / productOfferingPrice entities, TMF620-nested or flat.
                  Plural key names (e.g. productOfferings) are also accepted.
                </span>
              </div>
              <div className="flex items-start gap-2">
                <FileSpreadsheet size={14} className="mt-0.5 shrink-0" />
                <span>
                  CSV/Excel: columns id, type, name, lifecycleStatus, validFrom,
                  validTo, specId, priceId, categoryId, childOfferingIds.
                </span>
              </div>
            </div>
          </>
        )}

        {result && (
          <div>
            <div className="flex items-center justify-between mb-2 gap-3">
              <span className="text-xs text-slate-500 font-[IBM_Plex_Mono] truncate max-w-[40%]">
                {fileName}
              </span>
              <div className="flex items-center gap-4 shrink-0">
                <button
                  onClick={() => exportCSV(result, fileName)}
                  className="flex items-center gap-1.5 text-xs text-slate-400 hover:text-slate-200 transition-colors"
                >
                  <Download size={13} />
                  CSV
                </button>
                <button
                  onClick={() => exportPDF(result, fileName)}
                  className="flex items-center gap-1.5 text-xs text-slate-400 hover:text-slate-200 transition-colors"
                >
                  <FileDown size={13} />
                  PDF
                </button>
                <button
                  onClick={reset}
                  className="flex items-center gap-1.5 text-xs text-slate-400 hover:text-slate-200 transition-colors"
                >
                  <RotateCcw size={13} />
                  New scan
                </button>
              </div>
            </div>

            <div className="rounded-xl border border-slate-700/60 bg-[#0D1420] px-6 py-6 flex flex-col sm:flex-row items-center gap-6">
              <Gauge score={result.score} />
              <div className="flex flex-wrap gap-3 flex-1 w-full">
                <StatCard label="Entities" value={result.summary.entitiesScanned} />
                <StatCard label="Critical" value={result.summary.critical} accent="#F87171" />
                <StatCard label="Warnings" value={result.summary.warning} accent="#FB923C" />
                <StatCard label="Info" value={result.summary.info} accent="#60A5FA" />
                <StatCard label="Offerings" value={result.summary.offerings} />
              </div>
            </div>

            <ReportRail issues={result.issues} />
          </div>
        )}

        <div className="mt-16 pt-6 border-t border-slate-800 text-xs text-slate-600 leading-relaxed">
          Runs entirely in your browser — no catalog data is uploaded
          anywhere. Cross-system checks against a live source (Hansen EPC,
          Amdocs, BRM) aren't in this build; they'd need a backend connector.
          <div className="mt-2">
            <a
              href="https://www.tmforum.org/open-digital-architecture/open-apis/product-catalog-management-api-TMF620/"
              target="_blank"
              rel="noopener noreferrer"
              className="text-blue-600 hover:underline"
            >
              View TMF 620 Product Catalog Management API Documentation
            </a>
          </div>
        </div>
      </div>
    </div>
  );
}

export default function CatalogComplianceChecker() {
  return (
    <ErrorBoundary>
      <CatalogComplianceCheckerInner />
    </ErrorBoundary>
  );
}



