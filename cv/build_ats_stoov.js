/**
 * ATS-Safe CV Builder — Tom van der Arend
 * Één kolom, geen tabellen, geen afbeeldingen, lineaire flow.
 * Veilig voor alle ATS-systemen (Workday, Greenhouse, Lever, SAP SuccessFactors).
 *
 * Gebruik: Pas DATA sectie aan op de vacature, dan: node build_ats_cv.js
 */

const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  LevelFormat, BorderStyle, WidthType, HeadingLevel
} = require('docx');
const fs = require('fs');

// ─── KLEUREN (ATS-safe: alleen tekst, geen achtergronden) ────────────────────
const C = {
  darkGreen: "1A3D2B",   // naam + sectiekoppen
  midGreen:  "2D6A4F",   // functietitels
  accent:    "52946B",   // datums
  black:     "111111",
  darkGray:  "2C2C2C",
  gray:      "4A4A4A",
  lightGray: "888888",
};

// ─── OUTPUT ───────────────────────────────────────────────────────────────────
const OUTPUT_PATH = '/home/user/Default/cv/CV_Tom_van_der_Arend_Interim_Ecommerce_Manager_Stoov_ATS.docx';

// ─── PAGINA (A4, ruime marges voor leesbaarheid) ──────────────────────────────
const PAGE_W  = 11906;
const PAGE_H  = 16838;
const MARGIN  = 820;

// ═══════════════════════════════════════════════════════════════════════════════
// DATA — pas aan per vacature
// ═══════════════════════════════════════════════════════════════════════════════
const DATA = {
  naam:      "Tom van der Arend",
  functie:   "Interim E-commerce Manager  |  Shopify · CRO · Piekseizoen",
  contact: {
    woonplaats: "Driebergen-Rijsenburg, Nederland",
    telefoon:   "06 18 65 98 96",
    email:      "tomvanderarend@hotmail.com",
    linkedin:   "linkedin.com/in/tomvanderarend",
    website:    "marketflow.nl",
  },

  profiel: [
    "Hands-on Interim E-commerce Manager met 20+ jaar ervaring in e-commerce, digitale transformatie en IT, waarvan 13 jaar als co-founder van een eigen D2C-merk op Shopify. Bouwde Sfeer voor jou van nul naar 7 Europese markten, met meerdere talen, valuta en betaalmethoden, en draaide elk jaar zelf het Q4-piekseizoen van Black Friday, Sinterklaas en Kerst.",
    "Stuurt op CVR, AOV, RPV en marge en begrijpt hoe paid en attributie daarop inwerken. Combineert data uit Shopify en GA4 met een CRO-testroadmap en scherpe prioritering van de Shopify-roadmap. Landt snel als meewerkend teamlead: stuurt externe developers en bureaus aan zonder formele lijn, documenteert werkwijzen en draagt kennis over. Werkt AI-gedreven met Claude en agents om repeterend werk te automatiseren.",
  ],

  werkervaring: [
    {
      titel:   "Founder & Interim E-commerce Consultant",
      org:     "Marketflow",
      periode: "Juni 2022 — Heden",
      intro:   "Strategisch advies en hands-on implementatie op het snijvlak van e-commerce en IT voor D2C-merken, MKB en scale-ups.",
      bullets: [
        { tekst: "Begeleiding van e-commerce-, integratie- en digitaliseringsprojecten van requirements tot go-live" },
        { tekst: "AI-workflows en automatiseringen gebouwd met Claude en agents voor rapportage, content en repeterende processen" },
        { tekst: "KPI-dashboards en vast rapportageritme opgezet richting directie en eigenaren" },
      ]
    },
    {
      titel:   "Co-founder & Operationeel Verantwoordelijke",
      org:     "Sfeer voor jou (D2C, Shopify)",
      periode: "Juni 2012 — November 2025",
      intro:   "Eigen D2C-merk in lifestyle producten; van nul naar 7 Europese markten via Shopify en internationale marketplaces. Bewust afgebouwd om volledig te focussen op interim/consultancy.",
      bullets: [
        { tekst: "200%+ omzetgroei per jaar door datagedreven optimalisatie van conversie, merchandising en klantbeleving", vet: "200%+ omzetgroei per jaar" },
        { tekst: "13 jaar eigen Shopify-omgeving end-to-end beheerd: thema, Liquid-aanpassingen, metavelden, apps en API-koppelingen" },
        { tekst: "Internationale uitrol naar 7 markten: lokalisatie van talen, valuta, betaalmethoden en verzending" },
        { tekst: "Jaarlijks het Q4-piekseizoen gedraaid: commerciële kalender, campagnes, voorraad en een stabiele site onder piekbelasting" },
        { tekst: "Traffic en retentie omgezet in omzet via SEO/SEA, e-mail/CRM en marketplace management" },
      ]
    },
    {
      titel:   "IT Strategie Consultant & Project Lead",
      org:     "Dijk Natural Collections",
      periode: "September 2022 — Juni 2024",
      intro:   "Interim voor een PE-backed scale-up in de decoratiebranche; directe rapportage aan de directie.",
      bullets: [
        { tekst: "300% omzetgroei binnen één jaar door heldere e-commerce platformkeuzes en kanaalstrategie", vet: "300% omzetgroei binnen één jaar" },
        { tekst: "Roadmap geprioriteerd en parallelle trajecten bewaakt: e-commerce platform, ERP-selectie en procesoptimalisaties" },
        { tekst: "Externe developers en leveranciers aangestuurd; 30% efficiencyverbetering via herontwerp van workflows en KPI-structuren" },
      ]
    },
    {
      titel:   "IT Manager / Transitie & Organisatieverandering",
      org:     "Accenture Nederland",
      periode: "Oktober 2005 — November 2011",
      intro:   null,
      bullets: [
        { tekst: "Leiding aan IT-afdeling van 25 FTE; IT-transities voor Fortune 500-klanten met rapportage op board-niveau" },
      ]
    },
    {
      titel:   "Projectleider / Service Manager (incl. traineeship)",
      org:     "IBM Nederland",
      periode: "Mei 2001 — Oktober 2005",
      intro:   null,
      bullets: [
        { tekst: "Internationale IT-outsourcingtrajecten; SLA-borging en directe rapportage aan C-level" },
      ]
    },
  ],

  vaardigheden: [
    "Shopify (thema, Liquid, metavelden, apps)", "Shopify API & integraties", "CRO & A/B-testen",
    "UX & shopping flow", "Merchandising & commerciële kalender", "GA4 & tracking",
    "KPI-sturing: CVR, AOV, RPV, marge", "Paid & attributie", "Internationale e-commerce (markten, talen, valuta, betaalmethoden)",
    "SEO/SEA", "CRM & retentie", "Marketplace management",
    "AI-automatisering (Claude, agents)", "Roadmap & prioritering", "Aansturing developers & bureaus",
    "Coachend leiderschap", "Documentatie & kennisoverdracht",
  ],

  opleiding: [
    {
      titel:   "drs. Business Administration",
      org:     "Erasmus Universiteit Rotterdam",
      periode: "1994 — 2001",
      intro:   "Specialisaties: Informatie Management & Financieel Management",
    },
  ],

  activiteiten: [
    {
      titel:   "Penningmeester",
      org:     "Hockeyclub Shinty, Driebergen",
      periode: "November 2025 — Heden",
      intro:   null,
    },
  ],

  talen: ["Nederlands — Moedertaal", "Engels — Vloeiend"],
};

// ═══════════════════════════════════════════════════════════════════════════════
// BUILDER
// ═══════════════════════════════════════════════════════════════════════════════

function p(children, opts = {}) { return new Paragraph({ children, ...opts }); }
function r(text, opts = {}) { return new TextRun({ text, font: "Calibri", ...opts }); }

function sectionHeading(text) {
  return p([
    r(text.toUpperCase(), { bold: true, size: 21, color: C.darkGreen, characterSpacing: 60 })
  ], {
    spacing: { before: 200, after: 70 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.darkGreen, space: 3 } }
  });
}

function jobHeader(titel, org, periode) {
  return [
    p([
      r(titel, { bold: true, size: 20, color: C.darkGray }),
      r("   |   ", { size: 18, color: C.lightGray }),
      r(org, { size: 18, color: C.gray }),
    ], { spacing: { before: 130, after: 10 } }),
    p([r(periode, { size: 16, color: C.accent, italics: true })],
      { spacing: { before: 0, after: 60 } }),
  ];
}

function bodyPara(text) {
  return p([r(text, { size: 19, color: C.gray })],
    { spacing: { before: 0, after: 60 } });
}

function bullet({ tekst, vet }) {
  const children = [];
  if (vet) {
    const i = tekst.indexOf(vet);
    if (i > 0) children.push(r(tekst.slice(0, i), { size: 19, color: C.gray }));
    children.push(r(vet, { size: 19, color: C.darkGray, bold: true }));
    const rest = tekst.slice(i + vet.length);
    if (rest) children.push(r(rest, { size: 19, color: C.gray }));
  } else {
    children.push(r(tekst, { size: 19, color: C.gray }));
  }
  return new Paragraph({
    numbering: { reference: "atsbullets", level: 0 },
    spacing: { before: 24, after: 24 },
    children
  });
}

// ── Contactblok ───────────────────────────────────────────────────────────────
const contactChildren = [
  // Naam
  p([r(DATA.naam, { bold: true, size: 36, color: C.darkGreen })],
    { spacing: { before: 0, after: 60 } }),
  // Functie
  p([r(DATA.functie, { size: 22, color: C.gray, italics: true })],
    { spacing: { before: 0, after: 120 } }),
  // Contact op één regel
  p([
    r(DATA.contact.woonplaats, { size: 18, color: C.gray }),
    r("   ·   ", { size: 18, color: C.lightGray }),
    r(DATA.contact.telefoon, { size: 18, color: C.gray }),
    r("   ·   ", { size: 18, color: C.lightGray }),
    r(DATA.contact.email, { size: 18, color: C.gray }),
  ], { spacing: { before: 0, after: 30 } }),
  // LinkedIn + website
  p([
    r(DATA.contact.linkedin, { size: 18, color: C.midGreen }),
    r("   ·   ", { size: 18, color: C.lightGray }),
    r(DATA.contact.website, { size: 18, color: C.gray }),
  ], {
    spacing: { before: 0, after: 0 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.darkGreen, space: 8 } }
  }),
];

// ── Profiel ───────────────────────────────────────────────────────────────────
const profielChildren = [
  sectionHeading("Profiel"),
  ...DATA.profiel.map(bodyPara),
];

// ── Werkervaring ──────────────────────────────────────────────────────────────
const werkervaringChildren = [sectionHeading("Werkervaring")];
for (const job of DATA.werkervaring) {
  werkervaringChildren.push(...jobHeader(job.titel, job.org, job.periode));
  if (job.intro) werkervaringChildren.push(bodyPara(job.intro));
  for (const b of job.bullets) werkervaringChildren.push(bullet(b));
}

// ── Vaardigheden ──────────────────────────────────────────────────────────────
// ATS-tip: vaardigheden als kommalijst in één alinea — beter geparst dan bullets
const vaardighedenChildren = [
  sectionHeading("Vaardigheden"),
  p([r(DATA.vaardigheden.join("   ·   "), { size: 19, color: C.gray })],
    { spacing: { before: 0, after: 60 } }),
  p([r("Talen: " + DATA.talen.join("   ·   "), { size: 19, color: C.gray })],
    { spacing: { before: 0, after: 60 } }),
];

// ── Opleiding & Activiteiten ──────────────────────────────────────────────────
const opleidingChildren = [sectionHeading("Opleiding & Activiteiten")];
for (const item of [...DATA.opleiding, ...DATA.activiteiten]) {
  opleidingChildren.push(p([
    r(item.titel, { bold: true, size: 20, color: C.darkGray }),
    r("   |   ", { size: 18, color: C.lightGray }),
    r(item.org, { size: 18, color: C.gray }),
    r("   |   ", { size: 18, color: C.lightGray }),
    r(item.periode, { size: 16, color: C.accent, italics: true }),
  ], { spacing: { before: 100, after: 20 } }));
  if (item.intro) opleidingChildren.push(bodyPara(item.intro));
}

// ── Document ──────────────────────────────────────────────────────────────────
const allChildren = [
  ...contactChildren,
  ...profielChildren,
  ...werkervaringChildren,
  ...vaardighedenChildren,
  ...opleidingChildren,
];

const doc = new Document({
  numbering: {
    config: [{
      reference: "atsbullets",
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: "\u2022",  // standaard bullet — ATS-safe
        alignment: AlignmentType.LEFT,
        style: {
          run: { font: "Calibri", size: 19, color: C.accent },
          paragraph: { indent: { left: 440, hanging: 240 } }
        }
      }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Calibri", size: 19, color: C.gray } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: 640, right: MARGIN, bottom: 640, left: MARGIN }
      }
    },
    children: allChildren
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUTPUT_PATH, buf);
  console.log(`ATS CV opgeslagen: ${OUTPUT_PATH}`);
}).catch(e => { console.error(e); process.exit(1); });
