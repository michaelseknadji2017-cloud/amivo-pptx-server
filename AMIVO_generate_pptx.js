// AMIVO — AMIVO_generate_pptx.js
// Serveur Express qui reçoit le JSON Claude via POST et retourne un PPTX en base64
// Déploiement : Railway / Render (Node.js gratuit)
// Make.com appelle POST /generate avec le JSON Claude dans le body
// Usage local : node AMIVO_generate_pptx.js

const express = require("express");
const pptxgen = require("pptxgenjs");
const https = require("https");

const app = express();
app.use(express.json({ limit: "10mb" }));

// ============================================================
// CONFIG AIRTABLE
// ============================================================
const AIRTABLE_TOKEN = process.env.AIRTABLE_TOKEN;
const AIRTABLE_BASE_ID = "apprubsfPnIKKCVd1";
const AIRTABLE_TABLE_ID = "tblXHhIcW0qnpa7HI";

// Fetch record from Airtable by session_id
function fetchFromAirtable(sessionId) {
  return new Promise((resolve, reject) => {
    const formula = encodeURIComponent(`{session_id}="${sessionId}"`);
    const options = {
      hostname: "api.airtable.com",
      path: `/v0/${AIRTABLE_BASE_ID}/${AIRTABLE_TABLE_ID}?filterByFormula=${formula}&maxRecords=1`,
      method: "GET",
      headers: {
        "Authorization": `Bearer ${AIRTABLE_TOKEN}`,
        "Content-Type": "application/json"
      }
    };
    const req = https.request(options, (res) => {
      let data = "";
      res.on("data", chunk => data += chunk);
      res.on("end", () => {
        try {
          const json = JSON.parse(data);
          if (json.records && json.records.length > 0) {
            resolve(json.records[0].fields);
          } else {
            reject(new Error(`No record found for session_id: ${sessionId}`));
          }
        } catch (e) {
          reject(new Error(`Airtable parse error: ${e.message}`));
        }
      });
    });
    req.on("error", reject);
    req.end();
  });
}

// ============================================================
// PALETTE — Le Procès
// ============================================================
const C = {
  ink: "1A1209", paper: "F2E8D5", paperDark: "E8D9BE",
  red: "C0392B", gold: "B8860B", grey: "888888", white: "FFFFFF"
};
const makeShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 3, angle: 135, opacity: 0.3 });

function addGoldBar(slide, y, h = 0.03) {
  slide.addShape("rect", { x: 0, y, w: 10, h, fill: { color: C.gold }, line: { color: C.gold } });
}
function addStamp(slide, text, x, y, w = 2.5, rotate = -8) {
  slide.addShape("rect", { x, y, w, h: 0.45, fill: { color: "000000", transparency: 100 }, line: { color: C.red, width: 3 } });
  slide.addText(text, { x, y, w, h: 0.45, fontSize: 10, bold: true, color: C.red, align: "center", valign: "middle", charSpacing: 3, rotate, margin: 0 });
}
function addFooter(slide, num, prenom, destination) {
  slide.addShape("rect", { x: 0, y: 5.3, w: 10, h: 0.325, fill: { color: C.ink }, line: { color: C.ink } });
  addGoldBar(slide, 5.3, 0.025);
  slide.addText(`AMIVO — LE PROCÈS — ${(prenom || "").toUpperCase()} / ${(destination || "").toUpperCase()}`, { x: 0.3, y: 5.32, w: 7, h: 0.28, fontSize: 7, color: C.grey, charSpacing: 2, margin: 0 });
  slide.addText(`${num} / 12`, { x: 8.5, y: 5.32, w: 1.2, h: 0.28, fontSize: 7, color: C.grey, align: "right", margin: 0 });
}

// ============================================================
// PARSING JSON CLAUDE
// Gère les formats : JSON pur, ```json...```, {"type":"text","text":"..."}
// ============================================================
function parseClaudeOutput(raw) {
  let str = typeof raw === "string" ? raw : JSON.stringify(raw);

  // Format {"type":"text","text":"```json...```"}
  try {
    const wrapper = JSON.parse(str);
    if (wrapper.type === "text" && wrapper.text) str = wrapper.text;
  } catch (_) {}

  // Retire les backticks ```json ... ```
  str = str.replace(/^```json\s*/m, "").replace(/```\s*$/m, "").trim();

  return JSON.parse(str);
}

// ============================================================
// NORMALISATION — convertit le format Claude premium → format serveur
// ============================================================
function normalizePayload(D, fields) {
  const N = { ...D };

  // Champs depuis Airtable
  N.prenom         = fields.prenom        || D.prenom        || "";
  N.destination    = fields.destination   || D.destination   || "";
  N.dates          = fields.dates         || D.dates         || "";
  N.nb_participants = fields.nb_participants || D.nb_participants || "";
  N.session_id     = fields.session_id    || D.session_id    || "";

  // Alias simples
  N.dossier_intro  = D.dossier_intro  || D.crime_principal || "";
  N.mandat_complet = D.mandat_complet || D.mandat_arret    || "";
  N.verdict_peine  = D.verdict_peine  || D.verdict         || "";
  N.mot_de_fin     = D.mot_de_fin     || D.mot_temoin      || "";

  // chefs_accusation[] → chef_01_*, chef_02_*, chef_03_*
  if (Array.isArray(D.chefs_accusation)) {
    D.chefs_accusation.forEach((c, i) => {
      if (i >= 3) return;
      const n = i + 1;
      N[`chef_0${n}_intitule`] = c.chef      || c.intitule  || "";
      N[`chef_0${n}_detail`]   = c.detail    || "";
      N[`chef_0${n}_aggravant`]= c.aggravante|| c.aggravant || "";
    });
  }

  // temoins_liste
  if (!N.temoins_liste) N.temoins_liste = "";

  // programme_j1[] → j1_h1…j1_h6, j1_a1…, j1_d1…
  N.prog_j1_titre = D.prog_j1_titre || "Programme Jour 1";
  if (Array.isArray(D.programme_j1)) {
    D.programme_j1.forEach((s, i) => {
      if (i >= 6) return;
      const n = i + 1;
      N[`j1_h${n}`] = s.heure    || s.h || "";
      N[`j1_a${n}`] = s.activite || s.a || "";
      N[`j1_d${n}`] = s.description || s.d || "";
    });
  }

  // programme_j2[] → j2_h1…j2_h6, j2_a1…, j2_d1…
  N.prog_j2_titre = D.prog_j2_titre || "Programme Jour 2";
  if (Array.isArray(D.programme_j2)) {
    D.programme_j2.forEach((s, i) => {
      if (i >= 6) return;
      const n = i + 1;
      N[`j2_h${n}`] = s.heure    || s.h || "";
      N[`j2_a${n}`] = s.activite || s.a || "";
      N[`j2_d${n}`] = s.description || s.d || "";
    });
  }

  // jeux[] → jeu1_nom, jeu1_principe, jeu1_role1_prenom…, jeu1_etape1…, jeu1_question1…
  if (Array.isArray(D.jeux)) {
    D.jeux.forEach((jeu, i) => {
      if (i >= 5) return;
      const n = i + 1;
      N[`jeu${n}_nom`]      = jeu.nom      || "";
      N[`jeu${n}_emoji`]    = jeu.emoji    || "";
      N[`jeu${n}_duree`]    = jeu.duree    || "";
      N[`jeu${n}_lieu`]     = jeu.lieu     || "";
      N[`jeu${n}_principe`] = jeu.principe || "";
      N[`jeu${n}_materiel`] = jeu.materiel || "";

      // roles : string → role1_prenom / role1_role
      const rolesRaw = jeu.roles || "";
      const roleParts = typeof rolesRaw === "string"
        ? rolesRaw.split(",").map(r => r.trim()).filter(Boolean)
        : [];
      roleParts.slice(0, 4).forEach((r, ri) => {
        const words = r.split(" ");
        N[`jeu${n}_role${ri+1}_prenom`] = words[0] || "";
        N[`jeu${n}_role${ri+1}_role`]   = words.slice(1).join(" ") || r;
      });

      // deroulement[] → etape1…etape4
      const etapes = Array.isArray(jeu.deroulement) ? jeu.deroulement : (Array.isArray(jeu.etapes) ? jeu.etapes : []);
      etapes.slice(0, 4).forEach((e, ei) => { N[`jeu${n}_etape${ei+1}`] = e || ""; });

      // questions[]
      if (Array.isArray(jeu.questions)) {
        jeu.questions.slice(0, 3).forEach((q, qi) => { N[`jeu${n}_question${qi+1}`] = q || ""; });
      }
    });
  }

  // budget_shein[] → acc1_*, accessoires_budget_total
  let total = 0;
  if (Array.isArray(D.budget_shein)) {
    D.budget_shein.forEach((item, i) => {
      if (i >= 6) return;
      const n = i + 1;
      N[`acc${n}_categorie`]  = item.categorie   || "";
      N[`acc${n}_shein`]      = item.mot_cle     || item.shein || "";
      N[`acc${n}_quantite`]   = item.quantite    || "1";
      N[`acc${n}_prix_total`] = item.prix        || item.prix_total || "";
      N[`acc${n}_jeu`]        = item.jeu_associe || item.jeu || "";
      N[`acc${n}_priorite`]   = item.priorite    || "";
      const p = parseFloat((item.prix || "0").replace(/[^0-9.]/g, ""));
      if (!isNaN(p)) total += p;
    });
  }
  if (!N.accessoires_budget_total && total > 0) N.accessoires_budget_total = `~${Math.round(total)}€`;
  if (!N.accessoires_budget_par_personne) N.accessoires_budget_par_personne = "";
  if (!N.planning_j30) N.planning_j30 = "Commander les accessoires";
  if (!N.planning_j7)  N.planning_j7  = "Imprimer les documents";
  if (!N.planning_j1)  N.planning_j1  = "Briefer les amis";

  return N;
}

// ============================================================
// GÉNÉRATION PPTX
// ============================================================
async function generatePPTX(D) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = `AMIVO — Le Procès — ${D.prenom} / ${D.destination}`;

  const footer = (s, n) => addFooter(s, n, D.prenom, D.destination);

  // ── SLIDE 1 — COUVERTURE ──────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.ink };
    addGoldBar(s, 0, 0.04);
    addGoldBar(s, 5.585, 0.04);
    for (let i = 0; i < 8; i++) s.addShape("line", { x: i * 1.4, y: 0, w: 0, h: 5.625, line: { color: C.gold, width: 0.3, transparency: 85 } });
    s.addText("TRIBUNAL CORRECTIONNEL SUPRÊME", { x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 9, color: C.grey, align: "center", charSpacing: 5, margin: 0 });
    s.addText(`Dossier N° ${D.session_id}`, { x: 0.5, y: 0.58, w: 9, h: 0.25, fontSize: 8, color: C.gold, align: "center", charSpacing: 2, margin: 0 });
    s.addText([
      { text: "L'AFFAIRE", options: { fontSize: 20, color: C.grey, charSpacing: 8, breakLine: true } },
      { text: (D.prenom || "").toUpperCase(), options: { fontSize: 68, color: C.paper, bold: true, charSpacing: 4, breakLine: true } },
    ], { x: 0.5, y: 0.9, w: 9, h: 2.7, align: "center", valign: "middle", margin: 0 });
    s.addText(D.titre_narratif, { x: 1, y: 3.5, w: 8, h: 0.5, fontSize: 13, color: C.paper, align: "center", italic: true, margin: 0 });
    s.addShape("rect", { x: 2.5, y: 4.1, w: 5, h: 0.55, fill: { color: C.red }, line: { color: C.red }, shadow: makeShadow() });
    s.addText(`dit « ${D.surnom_officiel} »`, { x: 2.5, y: 4.1, w: 5, h: 0.55, fontSize: 11, color: C.white, align: "center", valign: "middle", bold: true, charSpacing: 1, margin: 0 });
    s.addText(`${D.destination}  ·  ${D.dates}  ·  ${D.nb_participants}`, { x: 1, y: 4.85, w: 8, h: 0.28, fontSize: 8, color: C.grey, align: "center", charSpacing: 2, margin: 0 });
    addStamp(s, "MISE EN EXAMEN", 7.0, 0.65, 2.8, -6);
    footer(s, "01");
  }

  // ── SLIDE 2 — DOSSIER D'ACCUSATION ───────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.ink };
    addGoldBar(s, 0, 0.04);
    s.addText("§ 01 — DOSSIER D'ACCUSATION", { x: 0.35, y: 0.1, w: 9, h: 0.38, fontSize: 9, color: C.grey, charSpacing: 4, align: "center", margin: 0 });
    s.addShape("rect", { x: 0.3, y: 0.6, w: 9.4, h: 1.05, fill: { color: "252015" }, line: { color: C.gold, width: 1 } });
    s.addText(D.dossier_intro, { x: 0.5, y: 0.67, w: 9, h: 0.92, fontSize: 10.5, color: C.paper, italic: true, valign: "middle", margin: 0 });

    const chefs = [
      { num: "01", titre: D.chef_01_intitule, detail: D.chef_01_detail, aggrav: D.chef_01_aggravant, x: 0.3 },
      { num: "02", titre: D.chef_02_intitule, detail: D.chef_02_detail, aggrav: D.chef_02_aggravant, x: 3.55 },
      { num: "03", titre: D.chef_03_intitule, detail: D.chef_03_detail, aggrav: D.chef_03_aggravant, x: 6.8 },
    ];
    chefs.forEach(({ num, titre, detail, aggrav, x }) => {
      s.addShape("rect", { x, y: 1.8, w: 3.1, h: 2.7, fill: { color: "1F180E" }, line: { color: C.red, width: 2 }, shadow: makeShadow() });
      s.addShape("rect", { x, y: 1.8, w: 3.1, h: 0.36, fill: { color: C.red }, line: { color: C.red } });
      s.addText(`Chef N° ${num}`, { x, y: 1.8, w: 3.1, h: 0.36, fontSize: 9, color: C.white, bold: true, align: "center", valign: "middle", charSpacing: 3, margin: 0 });
      s.addText(titre, { x: x + 0.1, y: 2.2, w: 2.9, h: 0.5, fontSize: 10, color: C.paper, bold: true, valign: "top", margin: 0 });
      s.addText(detail, { x: x + 0.1, y: 2.74, w: 2.9, h: 1.0, fontSize: 8.5, color: C.grey, valign: "top", margin: 0 });
      s.addText(`⚠ ${aggrav}`, { x: x + 0.1, y: 3.78, w: 2.9, h: 0.52, fontSize: 8, color: C.red, italic: true, valign: "top", margin: 0 });
    });

    s.addText("TÉMOINS À CHARGE :", { x: 0.3, y: 4.62, w: 2.5, h: 0.25, fontSize: 8, color: C.gold, bold: true, charSpacing: 2, margin: 0 });
    s.addText(D.temoins_liste, { x: 0.3, y: 4.87, w: 9.4, h: 0.25, fontSize: 9, color: C.paper, margin: 0 });
    footer(s, "02");
  }

  // ── SLIDE 3 — MANDAT D'ARRÊT ──────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: "F5EDD8" };
    addGoldBar(s, 0, 0.04);
    s.addText("RÉPUBLIQUE FRANÇAISE", { x: 0.5, y: 0.15, w: 9, h: 0.28, fontSize: 8, color: C.ink, align: "center", charSpacing: 5, bold: true, margin: 0 });
    s.addText("TRIBUNAL CORRECTIONNEL SUPRÊME", { x: 0.5, y: 0.4, w: 9, h: 0.26, fontSize: 8, color: C.ink, align: "center", charSpacing: 2, margin: 0 });
    s.addShape("line", { x: 0.5, y: 0.68, w: 9, h: 0, line: { color: C.ink, width: 2 } });
    s.addText("MANDAT D'ARRÊT", { x: 0.5, y: 0.76, w: 9, h: 0.55, fontSize: 26, color: C.ink, align: "center", bold: true, charSpacing: 6, margin: 0 });
    s.addText(`N° ${D.session_id}`, { x: 0.5, y: 1.28, w: 9, h: 0.25, fontSize: 8.5, color: C.grey, align: "center", italic: true, margin: 0 });
    s.addShape("line", { x: 0.5, y: 1.55, w: 9, h: 0, line: { color: C.ink, width: 1 } });
    s.addText(D.mandat_complet, { x: 0.6, y: 1.65, w: 8.8, h: 3.0, fontSize: 10, color: C.ink, valign: "top", lineSpacingMultiple: 1.35, margin: 0 });
    s.addShape("line", { x: 6.2, y: 4.75, w: 3.2, h: 0, line: { color: C.ink, width: 1 } });
    s.addText(`Le Procureur\n${D.temoins_liste ? D.temoins_liste.split(" · ")[0] : "Le Procureur"}`, { x: 6.2, y: 4.8, w: 3.2, h: 0.45, fontSize: 8.5, color: C.ink, align: "center", margin: 0 });
    addStamp(s, "OFFICIEL", 0.25, 0.25, 2.0, -12);
    addStamp(s, "À IMPRIMER", 0.25, 4.88, 2.2, -5);
    footer(s, "03");
  }

  // ── SLIDES 4-5 — PROGRAMME J1 & J2 ───────────────────────
  const progDays = [
    { num: "01", couleur: C.red, titre: D.prog_j1_titre, date: D.dates ? D.dates.split("—")[0]?.trim() || "Jour 1" : "Jour 1",
      slots: [
        { h: D.j1_h1, a: D.j1_a1, d: D.j1_d1 }, { h: D.j1_h2, a: D.j1_a2, d: D.j1_d2 },
        { h: D.j1_h3, a: D.j1_a3, d: D.j1_d3 }, { h: D.j1_h4, a: D.j1_a4, d: D.j1_d4 },
        { h: D.j1_h5, a: D.j1_a5, d: D.j1_d5 }, { h: D.j1_h6, a: D.j1_a6, d: D.j1_d6 },
      ], slideNum: "04"
    },
    { num: "02", couleur: C.gold, titre: D.prog_j2_titre, date: D.dates ? D.dates.split("—")[1]?.trim() || "Jour 2" : "Jour 2",
      slots: [
        { h: D.j2_h1, a: D.j2_a1, d: D.j2_d1 }, { h: D.j2_h2, a: D.j2_a2, d: D.j2_d2 },
        { h: D.j2_h3, a: D.j2_a3, d: D.j2_d3 }, { h: D.j2_h4, a: D.j2_a4, d: D.j2_d4 },
        { h: D.j2_h5, a: D.j2_a5, d: D.j2_d5 }, { h: D.j2_h6, a: D.j2_a6, d: D.j2_d6 },
      ], slideNum: "05"
    }
  ];

  progDays.forEach(({ num, couleur, titre, date, slots, slideNum }) => {
    const s = pres.addSlide();
    s.background = { color: C.ink };
    addGoldBar(s, 0, 0.04);
    s.addShape("rect", { x: 0, y: 0.04, w: 2.3, h: 0.7, fill: { color: couleur }, line: { color: couleur } });
    s.addText(`JOUR ${num}`, { x: 0, y: 0.04, w: 2.3, h: 0.7, fontSize: 18, color: num === "01" ? C.white : C.ink, bold: true, align: "center", valign: "middle", charSpacing: 4, margin: 0 });
    s.addText(titre || "", { x: 2.5, y: 0.1, w: 5.5, h: 0.6, fontSize: 18, color: C.paper, bold: true, valign: "middle", margin: 0 });
    s.addText(date, { x: 8.1, y: 0.15, w: 1.7, h: 0.5, fontSize: 8, color: C.gold, align: "right", valign: "middle", margin: 0 });
    addGoldBar(s, 0.74, 0.025);
    slots.forEach((slot, i) => {
      if (!slot.h) return;
      const y = 0.9 + i * 0.73;
      if (i > 0) s.addShape("line", { x: 0.3, y, w: 9.4, h: 0, line: { color: "2A2010", width: 0.5 } });
      s.addText(slot.h, { x: 0.3, y: y + 0.08, w: 0.85, h: 0.28, fontSize: 13, color: C.gold, bold: true, margin: 0 });
      s.addShape("oval", { x: 1.25, y: y + 0.16, w: 0.1, h: 0.1, fill: { color: C.red }, line: { color: C.red } });
      s.addText(slot.a || "", { x: 1.45, y: y + 0.06, w: 5.5, h: 0.28, fontSize: 11.5, color: C.paper, bold: true, margin: 0 });
      s.addText(slot.d || "", { x: 1.45, y: y + 0.34, w: 7.8, h: 0.26, fontSize: 8.5, color: C.grey, margin: 0 });
    });
    footer(s, slideNum);
  });

  // ── SLIDES 6-10 — LES 5 JEUX ─────────────────────────────
  const jeux = [1,2,3,4,5].map(n => ({
    n,
    nom: D[`jeu${n}_nom`], emoji: D[`jeu${n}_emoji`],
    duree: D[`jeu${n}_duree`], lieu: D[`jeu${n}_lieu`],
    principe: D[`jeu${n}_principe`],
    roles: [1,2,3,4].map(r => ({ p: D[`jeu${n}_role${r}_prenom`], r: D[`jeu${n}_role${r}_role`] })),
    etapes: [1,2,3,4].map(e => D[`jeu${n}_etape${e}`]),
    questions: [1,2,3].map(q => D[`jeu${n}_question${q}`]),
    materiel: D[`jeu${n}_materiel`],
    slide: String(n + 5).padStart(2, "0")
  }));

  jeux.forEach(({ n, nom, emoji, duree, lieu, principe, roles, etapes, questions, materiel, slide }) => {
    if (!nom) return;
    const s = pres.addSlide();
    s.background = { color: C.ink };
    addGoldBar(s, 0, 0.04);
    s.addShape("rect", { x: 0, y: 0.04, w: 0.58, h: 0.7, fill: { color: C.red }, line: { color: C.red } });
    s.addText(`${n}`, { x: 0, y: 0.04, w: 0.58, h: 0.7, fontSize: 22, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(`${emoji || ""}  ${nom}`, { x: 0.72, y: 0.1, w: 6.8, h: 0.6, fontSize: 20, color: C.paper, bold: true, valign: "middle", margin: 0 });
    s.addText(`⏱ ${duree || "?"} min  ·  ${lieu || ""}`, { x: 7.6, y: 0.15, w: 2.2, h: 0.5, fontSize: 8, color: C.gold, align: "right", margin: 0 });
    addGoldBar(s, 0.74, 0.025);

    s.addText("PRINCIPE", { x: 0.3, y: 0.85, w: 4.5, h: 0.26, fontSize: 8, color: C.gold, bold: true, charSpacing: 3, margin: 0 });
    s.addText(principe || "", { x: 0.3, y: 1.08, w: 4.5, h: 1.1, fontSize: 9.5, color: C.paper, valign: "top", margin: 0 });
    s.addText("RÔLES", { x: 0.3, y: 2.28, w: 4.5, h: 0.26, fontSize: 8, color: C.gold, bold: true, charSpacing: 3, margin: 0 });
    roles.forEach((role, i) => {
      if (!role.p) return;
      s.addShape("rect", { x: 0.3, y: 2.55 + i * 0.5, w: 4.4, h: 0.43, fill: { color: "1F180E" }, line: { color: "2A2010" } });
      s.addText(role.p, { x: 0.4, y: 2.57 + i * 0.5, w: 1.1, h: 0.39, fontSize: 9, color: C.red, bold: true, valign: "middle", margin: 0 });
      s.addText(role.r || "", { x: 1.55, y: 2.57 + i * 0.5, w: 3.1, h: 0.39, fontSize: 8.5, color: C.grey, valign: "middle", margin: 0 });
    });

    s.addText("DÉROULÉ", { x: 5.1, y: 0.85, w: 4.6, h: 0.26, fontSize: 8, color: C.gold, bold: true, charSpacing: 3, margin: 0 });
    etapes.forEach((e, i) => {
      if (!e) return;
      s.addShape("oval", { x: 5.1, y: 1.06 + i * 0.52, w: 0.22, h: 0.22, fill: { color: C.red }, line: { color: C.red } });
      s.addText(`${i+1}`, { x: 5.1, y: 1.05 + i * 0.52, w: 0.22, h: 0.22, fontSize: 7, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
      s.addText(e, { x: 5.4, y: 1.06 + i * 0.52, w: 4.3, h: 0.45, fontSize: 8.5, color: C.paper, valign: "top", margin: 0 });
    });
    s.addText("QUESTIONS PRÉPARÉES", { x: 5.1, y: 3.06, w: 4.6, h: 0.26, fontSize: 8, color: C.gold, bold: true, charSpacing: 3, margin: 0 });
    questions.forEach((q, i) => {
      if (!q) return;
      s.addText([
        { text: `Q${i+1}  `, options: { color: C.red, bold: true } },
        { text: q, options: { color: C.grey } }
      ], { x: 5.1, y: 3.3 + i * 0.5, w: 4.6, h: 0.44, fontSize: 8.5, valign: "top", margin: 0 });
    });

    s.addText("MATÉRIEL", { x: 0.3, y: 4.65, w: 2, h: 0.24, fontSize: 8, color: C.gold, bold: true, charSpacing: 3, margin: 0 });
    s.addText(materiel || "", { x: 0.3, y: 4.88, w: 9.4, h: 0.26, fontSize: 8.5, color: C.grey, margin: 0 });
    footer(s, slide);
  });

  // ── SLIDE 11 — BUDGET & ACCESSOIRES ──────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.ink };
    addGoldBar(s, 0, 0.04);
    s.addText("§ BUDGET & ACCESSOIRES SHEIN", { x: 0.35, y: 0.1, w: 9, h: 0.38, fontSize: 9, color: C.grey, charSpacing: 4, align: "center", margin: 0 });
    s.addShape("rect", { x: 0.3, y: 0.6, w: 9.4, h: 0.75, fill: { color: "1F180E" }, line: { color: C.gold, width: 2 }, shadow: makeShadow() });
    s.addText("BUDGET TOTAL ESTIMÉ", { x: 0.5, y: 0.67, w: 5, h: 0.3, fontSize: 9, color: C.grey, bold: true, charSpacing: 3, margin: 0 });
    s.addText(D.accessoires_budget_total || "~80€", { x: 6.8, y: 0.6, w: 2.7, h: 0.75, fontSize: 30, color: C.gold, bold: true, align: "right", valign: "middle", margin: 0 });
    s.addText(`soit ${D.accessoires_budget_par_personne || "~13€"} / personne`, { x: 0.5, y: 0.92, w: 5, h: 0.28, fontSize: 8.5, color: C.grey, margin: 0 });

    const cols = [["ARTICLE", 3.4], ["MOT-CLÉ SHEIN", 1.9], ["QTÉ", 0.55], ["PRIX", 0.75], ["JEU", 1.7], ["★★★", 0.7]];
    let xp = 0.3;
    cols.forEach(([label, w]) => {
      s.addShape("rect", { x: xp, y: 1.48, w, h: 0.28, fill: { color: C.red }, line: { color: C.red } });
      s.addText(label, { x: xp + 0.04, y: 1.48, w: w - 0.04, h: 0.28, fontSize: 7, color: C.white, bold: true, valign: "middle", charSpacing: 1, margin: 0 });
      xp += w;
    });

    const accs = [1,2,3,4,5,6].map(i => [
      D[`acc${i}_categorie`], D[`acc${i}_shein`], D[`acc${i}_quantite`],
      D[`acc${i}_prix_total`], D[`acc${i}_jeu`], D[`acc${i}_priorite`]
    ]).filter(r => r[0]);

    const cws = [3.4, 1.9, 0.55, 0.75, 1.7, 0.7];
    accs.forEach((row, i) => {
      const y = 1.76 + i * 0.44;
      const bg = i % 2 === 0 ? "1A140A" : "161008";
      s.addShape("rect", { x: 0.3, y, w: 9.4, h: 0.41, fill: { color: bg }, line: { color: "2A2010" } });
      let xx = 0.3;
      row.forEach((val, j) => {
        s.addText(val || "", { x: xx + 0.06, y, w: cws[j] - 0.06, h: 0.41, fontSize: j === 0 ? 9 : 8, color: j === 3 || j === 5 ? C.gold : j === 0 ? C.paper : C.grey, valign: "middle", margin: 0 });
        xx += cws[j];
      });
    });

    s.addText("PLANNING COMMANDE :", { x: 0.3, y: 4.6, w: 9.4, h: 0.24, fontSize: 8, color: C.gold, bold: true, charSpacing: 3, margin: 0 });
    s.addText([
      { text: "J-30 : ", options: { bold: true, color: C.red } }, { text: `${D.planning_j30 || "Commander sur Shein"}    `, options: { color: C.grey } },
      { text: "J-7 : ", options: { bold: true, color: C.gold } }, { text: `${D.planning_j7 || "Imprimer les documents"}    `, options: { color: C.grey } },
      { text: "J-1 : ", options: { bold: true, color: C.paper } }, { text: D.planning_j1 || "Briefer les amis", options: { color: C.grey } },
    ], { x: 0.3, y: 4.84, w: 9.4, h: 0.28, fontSize: 8.5, margin: 0 });
    footer(s, "11");
  }

  // ── SLIDE 12 — VERDICT FINAL ──────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.ink };
    addGoldBar(s, 0, 0.04);
    addGoldBar(s, 5.585, 0.04);
    s.addText("VERDICT FINAL", { x: 0.5, y: 0.18, w: 9, h: 0.36, fontSize: 9, color: C.grey, align: "center", charSpacing: 6, margin: 0 });
    s.addShape("rect", { x: 1.2, y: 0.65, w: 7.6, h: 1.1, fill: { color: "000000", transparency: 100 }, line: { color: C.red, width: 4 } });
    s.addText("COUPABLE", { x: 1.2, y: 0.65, w: 7.6, h: 1.1, fontSize: 54, color: C.red, bold: true, align: "center", valign: "middle", charSpacing: 10, margin: 0 });
    s.addText("Condamné à :", { x: 0.5, y: 1.88, w: 9, h: 0.26, fontSize: 9, color: C.grey, align: "center", margin: 0 });
    s.addText(D.verdict_peine || "", { x: 0.5, y: 2.12, w: 9, h: 0.5, fontSize: 14, color: C.paper, bold: true, align: "center", italic: true, margin: 0 });
    addGoldBar(s, 2.72, 0.025);
    s.addText("MOT DES TÉMOINS", { x: 0.5, y: 2.82, w: 9, h: 0.26, fontSize: 8, color: C.gold, bold: true, align: "center", charSpacing: 3, margin: 0 });
    s.addText(D.mot_de_fin || "", { x: 0.8, y: 3.1, w: 8.4, h: 1.6, fontSize: 11, color: C.paper, align: "center", italic: true, valign: "middle", margin: 0 });
    addGoldBar(s, 4.78, 0.025);
    s.addText("AMIVO", { x: 0.5, y: 4.88, w: 9, h: 0.36, fontSize: 9, color: C.grey, align: "center", charSpacing: 8, margin: 0 });
    s.addText("EVG & EVJF ULTRA-PERSONNALISÉS PAR IA", { x: 0.5, y: 5.2, w: 9, h: 0.22, fontSize: 7, color: "2A2A2A", align: "center", charSpacing: 3, margin: 0 });
    addStamp(s, "AUDIENCE LEVÉE", 6.7, 4.45, 3.0, -5);
  }

  // Retourne le PPTX en base64
  const base64 = await pres.write({ outputType: "base64" });
  return base64;
}

// ============================================================
// ROUTES EXPRESS
// ============================================================

// Health check
app.get("/", (req, res) => res.json({ status: "AMIVO PPTX Generator — OK" }));

// Route principale : POST /generate
// Body : { "session_id": "amivo-XXXX" } OU { "claude_output": "<json brut>" }
app.post("/generate", async (req, res) => {
  try {
    const { session_id, claude_output } = req.body;
    let DATA;

    if (session_id) {
      // Mode production : on fetch depuis Airtable
      console.log(`🔍 Fetching Airtable for session_id: ${session_id}`);
      const fields = await fetchFromAirtable(session_id);
      const rawJson = fields.payload_json;
      if (!rawJson) return res.status(400).json({ error: "No payload_json in Airtable record" });
      DATA = normalizePayload(parseClaudeOutput(rawJson), fields);
    } else if (claude_output) {
      // Mode test : JSON passé directement
      DATA = parseClaudeOutput(claude_output);
    } else {
      return res.status(400).json({ error: "Missing session_id or claude_output" });
    }

    console.log(`✅ Génération PPTX pour ${DATA.prenom} / ${DATA.destination}`);

    const base64 = await generatePPTX(DATA);

    res.json({
      success: true,
      prenom: DATA.prenom,
      filename: `AMIVO_${DATA.prenom}_${DATA.destination}_LeProcès.pptx`,
      pptx_base64: base64
    });
  } catch (err) {
    console.error("❌ Erreur génération PPTX:", err.message);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 AMIVO PPTX Server — port ${PORT}`));
