#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const PptxGenJS = require("pptxgenjs");

const ROOT = process.cwd();
const INPUTS = path.join(ROOT, "inputs");
const OUTPUTS = path.join(ROOT, "outputs");

function readJson(filePath, fallback) {
  try {
    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  } catch {
    return fallback;
  }
}

function extractData(payload) {
  if (payload && typeof payload === "object" && payload.data && typeof payload.data === "object") {
    return payload.data;
  }
  return payload || {};
}

function dateStamp() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(
    d.getDate()
  ).padStart(2, "0")}`;
}

function slugify(value) {
  return (value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "customer";
}

function parseArgs(argv) {
  const args = { period: "Current reporting period", customer: "Ventia", engagement: "Supply Chain" };
  if (argv[0] && !argv[0].startsWith("--")) {
    args.period = argv[0];
  }
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i] === "--period" && argv[i + 1]) args.period = argv[i + 1];
    if (argv[i] === "--customer" && argv[i + 1]) args.customer = argv[i + 1];
    if (argv[i] === "--engagement" && argv[i + 1]) args.engagement = argv[i + 1];
  }
  return args;
}

function normalize() {
  const slack = extractData(readJson(path.join(INPUTS, "slack.json"), { messages: [] }));
  const salesforce = extractData(readJson(path.join(INPUTS, "salesforce.json"), { opportunities: [] }));
  const gdrive = extractData(readJson(path.join(INPUTS, "gdrive.json"), { documents: [] }));
  const glean = extractData(readJson(path.join(INPUTS, "glean.json"), { insights: [] }));

  const actions = [];
  const risks = [];
  const accomplishments = [];
  const nextSteps = [];

  (slack.messages || []).forEach((m) => {
    actions.push({
      description: m.title || "Slack update",
      owner: m.owner || "Unknown",
      status: m.status || "In Progress",
      comments: m.action || "",
    });
    accomplishments.push(m.detail || m.title || "Slack update");
    if ((m.impact || "").toLowerCase() === "high" || (m.impact || "").toLowerCase() === "critical") {
      risks.push({
        description: `${m.title || "Slack risk"} - ${m.detail || ""}`.trim(),
        impact: m.impact || "High",
        probability: "Med",
        action: (m.action || "Mitigation TBD") + ` - ${m.owner || "Owner TBD"}`,
        status: m.status || "In Progress",
      });
    }
    if (m.action) nextSteps.push(m.action);
  });

  (salesforce.opportunities || []).forEach((o) => {
    actions.push({
      description: `${o.account || "Account"} - ${o.name || "Opportunity"}`,
      owner: o.owner || "Unknown",
      status: o.stage || "In Progress",
      comments: o.next_step || "",
    });
    if (o.detail) accomplishments.push(o.detail);
    if ((o.risk || "").toLowerCase() === "high" || (o.risk || "").toLowerCase() === "critical") {
      risks.push({
        description: `${o.name || "Opportunity"} - ${o.detail || ""}`.trim(),
        impact: o.risk || "High",
        probability: "Med",
        action: (o.next_step || "Mitigation TBD") + ` - ${o.owner || "Owner TBD"}`,
        status: o.stage || "In Progress",
      });
    }
    if (o.next_step) nextSteps.push(o.next_step);
  });

  (gdrive.documents || []).forEach((d) => {
    actions.push({
      description: d.title || "Document update",
      owner: d.owner || "Unknown",
      status: d.state || "In Progress",
      comments: d.required_action || "",
    });
    if (d.summary) accomplishments.push(d.summary);
    if (d.required_action) nextSteps.push(d.required_action);
  });

  (glean.insights || []).forEach((i) => {
    actions.push({
      description: i.topic || "Knowledge signal",
      owner: i.owner || "Unknown",
      status: i.state || "In Progress",
      comments: i.follow_up || "",
    });
    if (i.summary) accomplishments.push(i.summary);
    if (i.follow_up) nextSteps.push(i.follow_up);
  });

  return {
    actions: actions.slice(0, 12),
    risks: risks.slice(0, 10),
    accomplishments: accomplishments.slice(0, 8),
    nextSteps: nextSteps.slice(0, 8),
  };
}

function ragFromRisks(risks) {
  if (risks.some((r) => (r.impact || "").toLowerCase() === "critical")) return "Red";
  if (risks.some((r) => (r.impact || "").toLowerCase() === "high")) return "Amber";
  return "Green";
}

function addHeader(slide, title, subtitle) {
  slide.addShape("rect", { x: 0, y: 0, w: 13.33, h: 0.65, fill: { color: "FF3621" }, line: { color: "FF3621" } });
  slide.addText(title, {
    x: 0.3,
    y: 0.12,
    w: 9.5,
    h: 0.35,
    color: "FFFFFF",
    bold: true,
    fontSize: 16,
    fontFace: "Barlow",
  });
  slide.addText(subtitle, {
    x: 9.8,
    y: 0.16,
    w: 3.2,
    h: 0.28,
    align: "right",
    color: "FFFFFF",
    fontSize: 10,
    fontFace: "Barlow",
  });
}

function addFooter(slide) {
  slide.addText("Databricks Professional Services", {
    x: 0.3,
    y: 7.1,
    w: 6.0,
    h: 0.2,
    color: "7A7A7A",
    fontSize: 8,
    fontFace: "Barlow",
  });
}

function run() {
  if (!fs.existsSync(OUTPUTS)) fs.mkdirSync(OUTPUTS, { recursive: true });
  const parsed = parseArgs(process.argv.slice(2));
  const period = parsed.period;
  const customerName = parsed.customer;
  const engagementName = parsed.engagement;
  const data = normalize();
  const overall = ragFromRisks(data.risks);
  const reportDate = dateStamp(); // YYYY-MM-DD
  const reportDateDisplay = new Date().toLocaleDateString("en-AU", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Databricks PS";
  pptx.company = "Databricks";
  pptx.subject = `${customerName} weekly status`;
  pptx.title = `${customerName} ${engagementName} - Databricks PS Status Report`;
  pptx.theme = {
    headFontFace: "Barlow",
    bodyFontFace: "Barlow",
    lang: "en-AU",
  };

  const BRAND = {
    navy: "1B3139",
    cyan: "00A8E1",
    lava: "FF3621",
    white: "FFFFFF",
    gray100: "F7F9FA",
    gray300: "D6DEE2",
    gray500: "6C7A80",
    text: "1B3139",
    green: "00A86B",
    amber: "F5A623",
    red: "D32F2F",
  };
  const TOTAL_SLIDES = 11;
  let slideNo = 0;

  function statusFill(s) {
    if (s === "Green") return BRAND.green;
    if (s === "Amber") return BRAND.amber;
    return BRAND.red;
  }

  function addChrome(slide, title, subtitle) {
    slideNo += 1;
    slide.background = { color: BRAND.white };
    slide.addShape("rect", { x: 0, y: 0, w: 13.33, h: 0.66, fill: { color: BRAND.navy }, line: { color: BRAND.navy } });
    slide.addShape("rect", { x: 0, y: 0.66, w: 13.33, h: 0.04, fill: { color: BRAND.cyan }, line: { color: BRAND.cyan } });
    slide.addText(title, {
      x: 0.34,
      y: 0.12,
      w: 8.7,
      h: 0.28,
      color: BRAND.white,
      bold: true,
      fontSize: 15,
      fontFace: "Barlow",
    });
    slide.addText(subtitle, {
      x: 8.9,
      y: 0.16,
      w: 4.0,
      h: 0.22,
      align: "right",
      color: BRAND.white,
      fontSize: 9,
      fontFace: "Barlow",
    });

    slide.addShape("line", { x: 0.0, y: 6.95, w: 13.33, h: 0, line: { color: BRAND.gray300, pt: 0.75 } });
    slide.addText("Databricks Professional Services", {
      x: 0.32,
      y: 7.02,
      w: 5.8,
      h: 0.18,
      color: BRAND.gray500,
      fontSize: 8,
      fontFace: "Barlow",
    });
    slide.addText(`${slideNo} of ${TOTAL_SLIDES}`, {
      x: 11.9,
      y: 7.02,
      w: 1.0,
      h: 0.18,
      align: "right",
      color: BRAND.gray500,
      fontSize: 8,
      fontFace: "Barlow",
    });
  }

  // Slide 1: Cover
  let s = pptx.addSlide();
  addChrome(s, `${customerName} ${engagementName} - PS Engagement`, `Weekly Status Report | ${reportDateDisplay}`);
  s.addText("Weekly Status Report", {
    x: 0.7,
    y: 2.1,
    w: 7.2,
    h: 0.6,
    bold: true,
    fontSize: 40,
    color: BRAND.text,
    fontFace: "Barlow",
  });
  s.addText(reportDateDisplay, {
    x: 0.72,
    y: 2.8,
    w: 4.5,
    h: 0.35,
    fontSize: 16,
    color: BRAND.gray500,
    fontFace: "Barlow",
  });
  s.addShape("rect", { x: 8.6, y: 1.8, w: 3.8, h: 2.6, fill: { color: BRAND.gray100 }, line: { color: BRAND.gray300 } });
  s.addText(`Project Status\n${overall}`, {
    x: 8.9,
    y: 2.15,
    w: 3.2,
    h: 1.2,
    align: "center",
    valign: "mid",
    bold: true,
    fontSize: 24,
    color: statusFill(overall),
    fontFace: "Barlow",
  });

  // Slide 2: Agenda
  s = pptx.addSlide();
  addChrome(s, "Agenda", `${customerName} ${engagementName} - ${period}`);
  s.addText("Agenda", { x: 0.7, y: 1.0, w: 2.0, h: 0.3, bold: true, fontSize: 24, color: BRAND.text, fontFace: "Barlow" });
  const agenda = ["Teams", "Status Updates / Issues, risks", "High Level Plan", "Resource Plan", "Key Points to Discuss"];
  s.addText(
    agenda.map((a) => ({ text: a, options: { bullet: { indent: 18 }, breakLine: true } })),
    { x: 0.9, y: 1.6, w: 8.4, h: 3.6, fontSize: 17, color: BRAND.text, fontFace: "Barlow" }
  );

  // Slide 3: Teams
  s = pptx.addSlide();
  addChrome(s, "1) Teams", `Week ending ${reportDateDisplay}`);
  s.addShape("rect", { x: 0.6, y: 1.0, w: 5.9, h: 0.32, fill: { color: BRAND.lava }, line: { color: BRAND.lava } });
  s.addText(customerName, { x: 0.78, y: 1.05, w: 4.8, h: 0.2, bold: true, fontSize: 12, color: BRAND.white, fontFace: "Barlow" });
  s.addText(
    "Customer sponsor\nEngineering manager\nData and insights stakeholders",
    { x: 0.75, y: 1.5, w: 5.6, h: 2.4, fontSize: 12, color: BRAND.text, fontFace: "Barlow" }
  );
  s.addShape("rect", { x: 6.85, y: 1.0, w: 5.9, h: 0.32, fill: { color: BRAND.lava }, line: { color: BRAND.lava } });
  s.addText("Databricks", { x: 7.03, y: 1.05, w: 4.8, h: 0.2, bold: true, fontSize: 12, color: BRAND.white, fontFace: "Barlow" });
  s.addText(
    "Amin Movahed, Resident Solutions Architect\nMuthu Srinivasan, Senior PM\nDelivery Engineers and Account Team",
    { x: 7.0, y: 1.5, w: 5.6, h: 2.4, fontSize: 12, color: BRAND.text, fontFace: "Barlow" }
  );

  // Slide 4: Action items
  s = pptx.addSlide();
  addChrome(s, "2) Status Updates / Issues, Risks", period);
  s.addText("Action Items", { x: 0.5, y: 0.95, w: 2.5, h: 0.3, bold: true, fontSize: 16, color: BRAND.text, fontFace: "Barlow" });
  const actionRows = [
    [
      { text: "S. No.", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "Date", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "Description", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "Owner", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "Status", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "Comments", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
    ],
  ];
  data.actions.forEach((a, i) => {
    actionRows.push([String(i + 1), reportDateDisplay, a.description, a.owner, a.status, a.comments || ""]);
  });
  if (data.actions.length === 0) actionRows.push(["1", reportDateDisplay, "No action items captured", "-", "-", "-"]);
  s.addTable(actionRows, {
    x: 0.4,
    y: 1.3,
    w: 12.5,
    colW: [0.6, 1.1, 4.8, 1.5, 1.4, 3.1],
    rowH: 0.32,
    border: { color: BRAND.gray300, pt: 0.5 },
    fontFace: "Barlow",
    fontSize: 9,
    valign: "middle",
    color: BRAND.text,
  });

  // Slide 5: Engagement status
  s = pptx.addSlide();
  addChrome(s, "Engagement Status", period);
  s.addText("Project Status", { x: 0.5, y: 0.95, w: 2.0, h: 0.3, bold: true, fontSize: 13, color: BRAND.text, fontFace: "Barlow" });
  s.addText(`Overall: ${overall}`, {
    x: 0.5,
    y: 1.3,
    w: 2.4,
    h: 0.45,
    align: "center",
    bold: true,
    fontSize: 14,
    color: BRAND.white,
    fill: { color: statusFill(overall) },
    fontFace: "Barlow",
  });
  s.addText(`Scope: ${overall}   |   Schedule: ${overall}   |   Make-It-Right: Green`, {
    x: 3.1,
    y: 1.38,
    w: 9.3,
    h: 0.28,
    fontSize: 11,
    color: BRAND.text,
    fontFace: "Barlow",
  });
  s.addText("Status Summary", { x: 0.5, y: 1.7, w: 3.0, h: 0.3, bold: true, fontSize: 14, color: "1B3139", fontFace: "Barlow" });
  s.addText(
    "Weekly PM data consolidated from Slack, Salesforce, GDrive and Glean. Focus remains on schedule dependencies, action closure and stakeholder alignment.",
    { x: 0.5, y: 2.05, w: 6.2, h: 1.2, fontSize: 11, color: "1B3139", fontFace: "Barlow" }
  );
  s.addText("Key Accomplishments", { x: 7.0, y: 1.7, w: 3.0, h: 0.3, bold: true, fontSize: 14, color: "1B3139", fontFace: "Barlow" });
  s.addText(
    data.accomplishments.length
      ? data.accomplishments.map((t) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: true } }))
      : "No accomplishments captured.",
    { x: 7.0, y: 2.05, w: 5.8, h: 2.1, fontSize: 10.5, color: "1B3139", fontFace: "Barlow" }
  );
  s.addText("Activities for next period", { x: 0.5, y: 4.0, w: 4.0, h: 0.3, bold: true, fontSize: 14, color: "1B3139", fontFace: "Barlow" });
  s.addText(
    data.nextSteps.length
      ? data.nextSteps.map((t) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: true } }))
      : "No next steps captured.",
    { x: 0.5, y: 4.35, w: 12.0, h: 1.9, fontSize: 10.5, color: "1B3139", fontFace: "Barlow" }
  );
  s.addText("Legend: Complete | In Progress | At Risk | Blocked | Not Started", {
    x: 0.5,
    y: 6.6,
    w: 8.0,
    h: 0.25,
    fontSize: 9,
    color: BRAND.gray500,
    fontFace: "Barlow",
  });

  // Slide 6: Risks
  s = pptx.addSlide();
  addChrome(s, "Risk & Issue", period);
  const riskRows = [
    [
      { text: "ID", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "Type", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "Description", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "Impact", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "Probability", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "Action(s) - Owner", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "Status", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
    ],
  ];
  data.risks.forEach((r, idx) => {
    riskRows.push([String(idx + 1).padStart(2, "0"), "Risk", r.description, r.impact, r.probability, r.action, r.status]);
  });
  if (data.risks.length === 0) {
    riskRows.push(["01", "Risk", "No high-severity risks captured", "Low", "Low", "Continue monitoring - PM", "Open"]);
  }
  s.addTable(riskRows, {
    x: 0.35,
    y: 1.0,
    w: 12.6,
    colW: [0.6, 0.8, 4.2, 1.0, 1.1, 3.2, 1.7],
    rowH: 0.36,
    border: { color: BRAND.gray300, pt: 0.5 },
    fontFace: "Barlow",
    fontSize: 9,
    valign: "middle",
    color: BRAND.text,
  });

  // Slide 7: Resource plan
  s = pptx.addSlide();
  addChrome(s, "Resource Plan", period);
  const resRows = [
    [
      { text: "Name", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "Role", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "Hours", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "19/1", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "26/1", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "2/2", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "9/2", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "16/2", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      { text: "23/2", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
    ],
    ["Delivery Lead", "Project Delivery", "128", "16", "16", "16", "16", "16", "16"],
    ["Data Engineer", "Engineering", "240", "40", "40", "40", "40", "40", "40"],
    ["PM", "Project Management", "64", "8", "8", "8", "8", "8", "8"],
    ["RSA", "Architecture / Advisory", "32", "8", "8", "8", "8", "-", "-"],
  ];
  s.addTable(resRows, {
    x: 0.5,
    y: 1.1,
    w: 12.3,
    colW: [1.7, 2.4, 0.9, 0.8, 0.8, 0.8, 0.8, 0.9, 0.9],
    rowH: 0.42,
    border: { color: BRAND.gray300, pt: 0.5 },
    fontFace: "Barlow",
    fontSize: 10,
    color: BRAND.text,
  });
  s.addText(
    "The above resource plan is indicative and subject to change based on discovery and delivery outcomes.",
    { x: 0.5, y: 4.05, w: 11.5, h: 0.3, fontSize: 10, color: BRAND.gray500, fontFace: "Barlow" }
  );

  // Slide 8: Plan tracking
  s = pptx.addSlide();
  addChrome(s, "Plan Tracking", period);
  s.addText(`${customerName} ${engagementName} plan - tracked in customer Jira`, {
    x: 0.8,
    y: 2.2,
    w: 10.5,
    h: 0.8,
    fontSize: 20,
    bold: true,
    color: BRAND.text,
    fontFace: "Barlow",
  });

  // Slide 9: Thank you (matches sample flow)
  s = pptx.addSlide();
  addChrome(s, "Thank You", period);
  s.addText("Thank you", {
    x: 4.8,
    y: 3.0,
    w: 3.7,
    h: 0.8,
    bold: true,
    align: "center",
    fontSize: 40,
    color: BRAND.text,
    fontFace: "Barlow",
  });

  // Slide 10: Engagement status snapshot (prior week)
  s = pptx.addSlide();
  addChrome(s, "Engagement Status", "Prior week snapshot");
  s.addText(`Status Summary (${period})`, {
    x: 0.7,
    y: 1.2,
    w: 4.0,
    h: 0.3,
    bold: true,
    fontSize: 14,
    color: BRAND.text,
    fontFace: "Barlow",
  });
  s.addText(
    "Overall status remains controlled. Risks are visible and mitigation owners are assigned. Delivery cadence continues through design, implementation and validation tasks.",
    { x: 0.7, y: 1.55, w: 11.5, h: 1.0, fontSize: 11, color: BRAND.text, fontFace: "Barlow" }
  );
  s.addText("Milestones / Phases / Deliverables", {
    x: 0.7,
    y: 2.8,
    w: 4.5,
    h: 0.3,
    bold: true,
    fontSize: 14,
    color: BRAND.text,
    fontFace: "Barlow",
  });
  s.addTable(
    [
      [
        { text: "Item", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
        { text: "Status", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
        { text: "Target Date", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      ],
      ["Discovery & Conceptual Design", "Complete", "02/02/2026"],
      ["Logical Design and Implementation Plan", "In Progress", "06/02/2026"],
      ["Taxonomy Work Package", "In Progress", "20/02/2026"],
    ],
    {
      x: 0.7,
      y: 3.2,
      w: 8.5,
      colW: [4.8, 1.6, 2.1],
      rowH: 0.36,
      border: { color: BRAND.gray300, pt: 0.5 },
      fontFace: "Barlow",
      fontSize: 10,
      color: BRAND.text,
    }
  );

  // Slide 11: Engagement status snapshot (earlier week)
  s = pptx.addSlide();
  addChrome(s, "Engagement Status", "Earlier week snapshot");
  s.addText("Project Status", { x: 0.5, y: 0.95, w: 2.0, h: 0.3, bold: true, fontSize: 13, color: BRAND.text, fontFace: "Barlow" });
  s.addText(`Overall: ${overall}`, {
    x: 0.5,
    y: 1.3,
    w: 2.4,
    h: 0.45,
    align: "center",
    bold: true,
    fontSize: 14,
    color: BRAND.white,
    fill: { color: statusFill(overall) },
    fontFace: "Barlow",
  });
  s.addText(`Scope: ${overall}   |   Schedule: ${overall}   |   Make-It-Right: Green`, {
    x: 3.1,
    y: 1.38,
    w: 9.3,
    h: 0.28,
    fontSize: 11,
    color: BRAND.text,
    fontFace: "Barlow",
  });
  s.addText("Points to discuss", {
    x: 0.5,
    y: 2.0,
    w: 3.0,
    h: 0.3,
    bold: true,
    fontSize: 14,
    color: BRAND.text,
    fontFace: "Barlow",
  });
  const points = data.risks.length
    ? data.risks.slice(0, 4).map((r) => `${r.description}`)
    : data.nextSteps.length
      ? data.nextSteps.slice(0, 4)
      : ["No priority discussion points captured for this period."];
  s.addText(points.map((p) => ({ text: p, options: { bullet: { indent: 12 }, breakLine: true } })), {
    x: 0.6,
    y: 2.35,
    w: 12.0,
    h: 1.55,
    fontSize: 11,
    color: BRAND.text,
    fontFace: "Barlow",
  });

  s.addText("Milestones / Phases / Deliverables", {
    x: 0.5,
    y: 4.15,
    w: 4.8,
    h: 0.3,
    bold: true,
    fontSize: 14,
    color: BRAND.text,
    fontFace: "Barlow",
  });
  s.addTable(
    [
      [
        { text: "Item", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
        { text: "Status", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
        { text: "Target Date", options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy } } },
      ],
      ["Discovery & Design", "Complete", "02/02/2026"],
      ["Implementation (Modelling, Engineering)", "In Progress", "06/02/2026"],
      ["Validation and Readout", "Not Started", "TBC"],
    ],
    {
      x: 0.5,
      y: 4.55,
      w: 8.8,
      colW: [5.0, 1.7, 2.1],
      rowH: 0.34,
      border: { color: BRAND.gray300, pt: 0.5 },
      fontFace: "Barlow",
      fontSize: 10,
      color: BRAND.text,
    }
  );

  s.addText("Legend: Complete | In Progress | At Risk | Blocked | Not Started", {
    x: 0.5,
    y: 6.58,
    w: 8.0,
    h: 0.25,
    fontSize: 9,
    color: BRAND.gray500,
    fontFace: "Barlow",
  });

  const outputName = `${slugify(customerName)}_${slugify(engagementName)}_databricks_ps_status_report_${reportDate}.pptx`;
  const outputPath = path.join(OUTPUTS, outputName);
  pptx
    .writeFile({ fileName: outputPath })
    .then(() => {
      console.log(`Generated PPTX: ${outputPath}`);
    })
    .catch((err) => {
      console.error("Failed to generate PPTX:", err);
      process.exitCode = 1;
    });
}

run();
