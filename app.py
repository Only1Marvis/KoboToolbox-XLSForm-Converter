# ── Imports and Setup ─────────────────────────────────────────────────────────
import os
import uuid
from flask import Flask, request, send_file, jsonify, Response
from convert import convert

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
ALLOWED_EXTENSIONS = {'docx', 'doc', 'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

REPEAT_GUIDE_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>The Enubiaka XLSForm Repeat Group Standards</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: linear-gradient(135deg, #f0f4ff 0%, #e8f0fe 100%);
            min-height: 100vh; padding: 2rem 1rem;
        }
        .container {
            background: white; border-radius: 16px; padding: 2.5rem;
            width: 100%; max-width: 900px; margin: 0 auto;
            box-shadow: 0 4px 24px rgba(0,0,0,0.08);
        }
        h1 { font-size: 1.7rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0.3rem; }
        .subtitle { color: #6b7280; font-size: 0.9rem; margin-bottom: 2rem; }
        .back-link {
            display: inline-block; margin-bottom: 1.5rem;
            color: #4472C4; text-decoration: none; font-size: 0.85rem; font-weight: 600;
        }
        .back-link:hover { text-decoration: underline; }
        .quick-ref {
            background: #f8faff; border: 1px solid #e0e7ff;
            border-radius: 10px; padding: 1.25rem; margin-bottom: 2rem;
        }
        .quick-ref h3 { font-size: 0.9rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0.75rem; }
        .quick-ref table { width: 100%; border-collapse: collapse; font-size: 0.82rem; }
        .quick-ref th { background: #4472C4; color: white; padding: 0.5rem 0.75rem; text-align: left; }
        .quick-ref td { padding: 0.45rem 0.75rem; border-bottom: 1px solid #e5e7eb; }
        .quick-ref tr:last-child td { border-bottom: none; }
        .quick-ref tr:nth-child(even) td { background: #f3f6ff; }
        .default-badge {
            background: #4472C4; color: white; font-size: 0.65rem;
            padding: 1px 6px; border-radius: 10px; font-weight: 700; margin-left: 4px;
        }
        .format-card {
            border: 1px solid #e5e7eb; border-radius: 12px;
            margin-bottom: 1.75rem; overflow: hidden;
        }
        .format-header {
            background: #f8faff; border-bottom: 1px solid #e5e7eb;
            padding: 1rem 1.25rem; display: flex; align-items: center; gap: 0.75rem;
        }
        .format-num {
            background: #4472C4; color: white; width: 28px; height: 28px;
            border-radius: 50%; display: flex; align-items: center; justify-content: center;
            font-size: 0.85rem; font-weight: 700; flex-shrink: 0;
        }
        .format-title { font-size: 1rem; font-weight: 700; color: #1a1a2e; }
        .format-body { padding: 1.25rem; }
        .what-it-is { font-size: 0.88rem; color: #374151; line-height: 1.7; margin-bottom: 1rem; }
        .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-bottom: 1rem; }
        .col-box { background: #f8faff; border-radius: 8px; padding: 0.85rem; }
        .col-box h4 { font-size: 0.78rem; font-weight: 700; color: #4472C4; margin-bottom: 0.5rem; text-transform: uppercase; letter-spacing: 0.05em; }
        .col-box ul { list-style: none; padding: 0; }
        .col-box ul li { font-size: 0.8rem; color: #374151; padding: 0.2rem 0; padding-left: 1rem; position: relative; line-height: 1.5; }
        .col-box ul li::before { content: "•"; position: absolute; left: 0; color: #4472C4; }
        .col-box.warning h4 { color: #b45309; }
        .col-box.warning ul li::before { color: #b45309; }
        .raw-data { background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; padding: 0.85rem; margin-top: 0.5rem; }
        .raw-data h4 { font-size: 0.78rem; font-weight: 700; color: #374151; margin-bottom: 0.4rem; text-transform: uppercase; letter-spacing: 0.05em; }
        .raw-data p { font-size: 0.8rem; color: #6b7280; line-height: 1.6; }
        .credit { text-align: center; margin-top: 2rem; font-size: 0.78rem; color: #9ca3af; }
        @media (max-width: 600px) { .two-col { grid-template-columns: 1fr; } }
    </style>
</head>
<body>
<div class="container">
    <a href="/" class="back-link">← Back to Converter</a>
    <h1>The Enubiaka XLSForm Repeat Group Standards</h1>
    <p class="subtitle">A reference guide for selecting the correct repeat group format for your survey design.</p>

    <!-- Quick Reference Table -->
    <div class="quick-ref">
        <h3>📋 Quick Reference Overview</h3>
        <table>
            <tr>
                <th>Format</th>
                <th>Structure</th>
                <th>Default?</th>
                <th>Best For</th>
            </tr>
            <tr>
                <td><strong>1. Sequential Positional Loop</strong></td>
                <td>Unrolled groups by slot position</td>
                <td>No</td>
                <td>Fixed-count sequential recording</td>
            </tr>
            <tr>
                <td><strong>2. Brand-Fixed Conditional Group Loop</strong></td>
                <td>One fixed group per brand/item</td>
                <td>No</td>
                <td>Brand-dedicated column analysis</td>
            </tr>
            <tr>
                <td>
                    <strong>3. Automated Selection-Driven Repeat Loop</strong>
                    <span class="default-badge">Default</span>
                </td>
                <td>True begin_repeat / end_repeat</td>
                <td>✅ Yes</td>
                <td>Variable-count, large lists, long format analysis</td>
            </tr>
        </table>
    </div>

    <!-- Format 1 -->
    <div class="format-card">
        <div class="format-header">
            <div class="format-num">1</div>
            <div class="format-title">Sequential Positional Loop</div>
        </div>
        <div class="format-body">
            <p class="what-it-is">A fixed number of unrolled groups (one per slot position) where the enumerator manually selects a brand or item at the start of each group via a select_one question. Already-selected items are excluded from subsequent groups via choice_filter, preventing duplicates. The total number of groups is fixed at design time.</p>
            <div class="two-col">
                <div class="col-box">
                    <h4>✅ When To Use</h4>
                    <ul>
                        <li>When enumerator control over item order is required</li>
                        <li>When total number of items is fixed in advance</li>
                        <li>When a guided one-item-at-a-time CAPI experience is needed</li>
                        <li>When the list is small (ideally 10 items or fewer)</li>
                    </ul>
                </div>
                <div class="col-box warning">
                    <h4>⚠ When NOT To Use</h4>
                    <ul>
                        <li>When the number of items varies per respondent</li>
                        <li>When brand-fixed columns in raw data are required</li>
                        <li>When the item list is large</li>
                        <li>When enumerator control over order is not needed</li>
                    </ul>
                </div>
            </div>
            <div class="col-box" style="margin-bottom:0.5rem;">
                <h4>📊 Best Suited Study Types</h4>
                <ul>
                    <li>Product testing studies — handing out multiple product samples sequentially</li>
                    <li>Sensory evaluation studies — taste tests, fragrance tests, texture evaluations</li>
                    <li>Mystery shopping audits — visiting a fixed number of outlets one at a time</li>
                    <li>Ranked preference studies — deliberate item presentation order</li>
                    <li>Experimental studies with controlled stimulus order</li>
                </ul>
            </div>
            <div class="raw-data">
                <h4>📁 Raw Data Appearance</h4>
                <p>Exports as a single flat sheet. Columns are organised by slot position, not by item. The same brand can appear in different slot columns across interviews. Cross-respondent analysis by brand requires filtering by slotN_pick value.</p>
            </div>
        </div>
    </div>

    <!-- Format 2 -->
    <div class="format-card">
        <div class="format-header">
            <div class="format-num">2</div>
            <div class="format-title">Brand-Fixed Conditional Group Loop</div>
        </div>
        <div class="format-body">
            <p class="what-it-is">Each brand or item has its own permanently fixed group with hardcoded questions and a unique field name prefix. The enumerator selects all applicable brands upfront via a select_multiple question, and only the selected brands' groups appear — always in the same fixed order. No reshaping is needed for brand-level analysis.</p>
            <div class="two-col">
                <div class="col-box">
                    <h4>✅ When To Use</h4>
                    <ul>
                        <li>When each brand must always occupy its own dedicated columns</li>
                        <li>When cross-respondent brand comparison is a priority</li>
                        <li>When the brand list is fully known, stable, and small (under 20 items)</li>
                        <li>When all brands can be selected upfront before questions begin</li>
                    </ul>
                </div>
                <div class="col-box warning">
                    <h4>⚠ When NOT To Use</h4>
                    <ul>
                        <li>When the brand list exceeds 20 items</li>
                        <li>When a guided one-at-a-time sequential flow is needed</li>
                        <li>When brands are dynamic or change between survey rounds</li>
                        <li>When respondents have vastly different numbers of applicable brands</li>
                    </ul>
                </div>
            </div>
            <div class="col-box" style="margin-bottom:0.5rem;">
                <h4>📊 Best Suited Study Types</h4>
                <ul>
                    <li>Brand equity and brand health tracking studies</li>
                    <li>Retail shelf audits — availability, pricing, and placement for a fixed set of SKUs</li>
                    <li>Competitive benchmarking studies</li>
                    <li>Media monitoring surveys — recall, sentiment and exposure for defined channels</li>
                    <li>Policy compliance audits and facility assessment surveys</li>
                </ul>
            </div>
            <div class="raw-data">
                <h4>📁 Raw Data Appearance</h4>
                <p>Exports as a single flat sheet with brand-dedicated columns. Each brand always appears under the same columns regardless of interview or selection order. Brands not selected are simply blank. No filtering or reshaping needed for brand-level analysis.</p>
            </div>
        </div>
    </div>

    <!-- Format 3 -->
    <div class="format-card" style="border-color: #4472C4;">
        <div class="format-header" style="background: #f0f4ff; border-bottom-color: #4472C4;">
            <div class="format-num">3</div>
            <div class="format-title">
                Automated Selection-Driven Repeat Loop
                <span class="default-badge">✅ Default</span>
            </div>
        </div>
        <div class="format-body">
            <p class="what-it-is">A true begin_repeat / end_repeat group driven by a prior select_multiple question. The repeat runs automatically — once per selected item — iterating through each selection in order using selected-at(). The number of iterations adapts to each respondent's selections. This is the recommended default for most survey designs.</p>
            <div class="two-col">
                <div class="col-box">
                    <h4>✅ When To Use</h4>
                    <ul>
                        <li>When the number of iterations varies per respondent</li>
                        <li>When the item list is large or may change between rounds</li>
                        <li>When enumerator control over item order is not required</li>
                        <li>When analysis will be done in long format or reshaped programmatically</li>
                        <li>When form performance and maintainability are priorities</li>
                    </ul>
                </div>
                <div class="col-box warning">
                    <h4>⚠ When NOT To Use</h4>
                    <ul>
                        <li>When brand-fixed columns are strictly required without reshaping</li>
                        <li>When enumerator control over which item appears next is required</li>
                        <li>When complex cross-iteration referencing or deduplication logic is needed</li>
                    </ul>
                </div>
            </div>
            <div class="col-box" style="margin-bottom:0.5rem;">
                <h4>📊 Best Suited Study Types</h4>
                <ul>
                    <li>Household surveys — details per household member, crop, asset, or income source</li>
                    <li>Health and nutrition studies — per child, illness episode, or food item</li>
                    <li>Agricultural surveys — per plot, crop type, or livestock category</li>
                    <li>Market surveys with large or variable brand lists</li>
                    <li>Panel and longitudinal studies with long-to-wide reshaping in Stata/R</li>
                    <li>Any study where item count is unknown at design time or varies widely</li>
                </ul>
            </div>
            <div class="raw-data">
                <h4>📁 Raw Data Appearance</h4>
                <p>Exports as two linked sheets — a main sheet (one row per submission) and a repeat sheet (one row per iteration per submission). All iterations share the same column names and are distinguished by the current_item value. Requires filtering by current_item or reshaping in Stata/R to analyse by specific item.</p>
            </div>
        </div>
    </div>

    <div class="credit">
        The Enubiaka XLSForm Repeat Group Standards &nbsp;|&nbsp; By Marvis Onyenwenu Enubiaka
    </div>
</div>
<script>
function applyPageLang(lang) {
  if (lang !== 'fr') return;
  var T = {
    title:   "Les normes de groupes répétés Enubiaka XLSForm",
    subtitle:"Guide de référence pour choisir le bon format de groupe répété pour votre conception d'enquête.",
    back:    "← Retour au convertisseur",
    qr_title:"📋 Vue d'ensemble rapide",
    qr_th:   ["Format","Structure","Par défaut ?","Idéal pour"],
    f1_title:"1. Boucle positionnelle séquentielle",
    f2_title:"2. Boucle de groupe fixe par marque",
    f3_title:"3. Boucle répétée automatisée pilotée par sélection",
    default_badge:"✅ Par défaut",
    when_use:"✅ Quand utiliser",
    when_not:"⚠ Quand NE PAS utiliser",
    best_for:"📊 Types d'études les mieux adaptés",
    raw_data:"📁 Apparence des données brutes",
    credit:  "Les normes de groupes répétés Enubiaka XLSForm | Par Marvis Onyenwenu Enubiaka",
    qr_rows: [
      ["<strong>1. Boucle positionnelle séquentielle</strong>","Groupes déroulés par position de slot","Non","Enregistrement séquentiel à nombre fixe"],
      ["<strong>2. Boucle de groupe fixe par marque</strong>","Un groupe fixe par marque/article","Non","Analyse en colonne dédiée par marque"],
      ["<strong>3. Boucle répétée automatisée</strong> <span class='default-badge'>Par défaut</span>","Vrai begin_repeat / end_repeat","✅ Oui","Listes variables/grandes, analyses en format long"]
    ],
    f1: {
      desc:"Un nombre fixe de groupes déroulés où l'enquêteur sélectionne manuellement une marque via select_one à chaque slot. Les articles déjà sélectionnés sont exclus via choice_filter. Le nombre total de groupes est fixé à la conception.",
      use:["Quand le contrôle de l'ordre des articles par l'enquêteur est requis","Quand le nombre total d'articles est fixé à l'avance","Quand une expérience CAPI guidée un-à-la-fois est nécessaire","Quand la liste est petite (idéalement 10 articles ou moins)"],
      not:["Quand le nombre d'articles varie selon le répondant","Quand des colonnes fixes par marque dans les données brutes sont requises","Quand la liste d'articles est grande","Quand le contrôle de l'ordre n'est pas nécessaire"],
      best:["Études de test produit — distribuer plusieurs échantillons séquentiellement","Études d'évaluation sensorielle — tests de goût, de parfum, de texture","Audits mystère — visiter un nombre fixe de points de vente","Études de préférence classée — ordre de présentation délibéré","Études expérimentales avec ordre de stimuli contrôlé"],
      raw:"Exporté en une feuille plate. Les colonnes sont organisées par position de slot. La même marque peut apparaître dans différentes colonnes selon les entretiens. L'analyse par marque nécessite un filtrage par la valeur slotN_pick."
    },
    f2: {
      desc:"Chaque marque possède son propre groupe permanent avec des questions fixes et un préfixe unique. L'enquêteur sélectionne toutes les marques en amont via select_multiple, et seuls les groupes des marques sélectionnées apparaissent — toujours dans le même ordre fixe.",
      use:["Quand chaque marque doit occuper ses propres colonnes dédiées","Quand la comparaison de marques entre répondants est prioritaire","Quand la liste de marques est connue, stable et petite (moins de 20)","Quand toutes les marques peuvent être sélectionnées en amont"],
      not:["Quand la liste de marques dépasse 20 articles","Quand un flux séquentiel guidé est nécessaire","Quand les marques sont dynamiques ou changent entre les vagues","Quand les répondants ont des nombres très différents de marques"],
      best:["Études de brand equity et de santé de marque","Audits de rayon — disponibilité, prix et placement pour un ensemble fixe de SKU","Études de benchmarking concurrentiel","Enquêtes de monitoring médias","Audits de conformité et enquêtes d'évaluation d'établissements"],
      raw:"Exporté en une feuille plate avec des colonnes dédiées par marque. Chaque marque apparaît toujours sous les mêmes colonnes. Les marques non sélectionnées sont vides. Aucun remodelage nécessaire pour l'analyse par marque."
    },
    f3: {
      desc:"Un vrai groupe begin_repeat / end_repeat piloté par une question select_multiple précédente. La répétition s'exécute automatiquement — une fois par article sélectionné — en utilisant selected-at(). Le nombre d'itérations s'adapte aux sélections de chaque répondant. Format par défaut recommandé.",
      use:["Quand le nombre d'itérations varie selon le répondant","Quand la liste d'articles est grande ou peut changer entre les vagues","Quand le contrôle de l'ordre n'est pas requis","Quand l'analyse sera faite en format long ou remodelée programmatiquement","Quand la performance et la maintenabilité du formulaire sont des priorités"],
      not:["Quand des colonnes fixes par marque sont strictement requises","Quand le contrôle de l'enquêteur sur l'article suivant est requis","Quand une logique complexe de référencement entre itérations est nécessaire"],
      best:["Enquêtes ménages — détails par membre, culture, actif ou source de revenus","Études de santé et nutrition — par enfant, épisode morbide ou aliment","Enquêtes agricoles — par parcelle, type de culture ou bétail","Enquêtes de marché avec grandes listes de marques","Études de panel et longitudinales avec remodelage dans Stata/R","Toute étude où le nombre d'articles est inconnu à la conception"],
      raw:"Exporté en deux feuilles liées — une feuille principale (une ligne par soumission) et une feuille de répétition (une ligne par itération). Toutes les itérations partagent les mêmes noms de colonnes. Nécessite un filtrage par current_item ou un remodelage dans Stata/R."
    }
  };

  var h1 = document.querySelector('h1'); if(h1) h1.textContent = T.title;
  var sub = document.querySelector('.subtitle'); if(sub) sub.textContent = T.subtitle;
  var bl = document.querySelector('.back-link'); if(bl) bl.textContent = T.back;

  // Quick ref
  var qrh = document.querySelector('.quick-ref h3'); if(qrh) qrh.textContent = T.qr_title;
  var qrths = document.querySelectorAll('.quick-ref th');
  T.qr_th.forEach(function(t,i){ if(qrths[i]) qrths[i].textContent = t; });
  var qrtrs = document.querySelectorAll('.quick-ref tr');
  T.qr_rows.forEach(function(row,ri){
    var tr = qrtrs[ri+1]; if(!tr) return;
    var tds = tr.querySelectorAll('td');
    row.forEach(function(c,ci){ if(tds[ci]) tds[ci].innerHTML = c; });
  });

  // Format cards
  var cards = document.querySelectorAll('.format-card');
  var formats = [T.f1, T.f2, T.f3];
  var ftitles = [T.f1_title, T.f2_title, T.f3_title];

  cards.forEach(function(card,idx){
    var fd = formats[idx]; if(!fd) return;
    var ft = card.querySelector('.format-title');
    if(ft) {
      ft.innerHTML = ftitles[idx];
      if(idx === 2) ft.innerHTML += ' <span class="default-badge">' + T.default_badge + '</span>';
    }
    var desc = card.querySelector('.what-it-is'); if(desc) desc.textContent = fd.desc;
    card.querySelectorAll('.col-box').forEach(function(col){
      var h4 = col.querySelector('h4'); if(!h4) return;
      var txt = h4.textContent.trim();
      var lis, items;
      if(txt.includes('When To Use')||txt.includes('✅')){
        h4.textContent = T.when_use; items = fd.use;
      } else if(txt.includes('When NOT')||txt.includes('⚠')){
        h4.textContent = T.when_not; items = fd.not;
      } else if(txt.includes('Best Suited')||txt.includes('📊')){
        h4.textContent = T.best_for; items = fd.best;
      }
      if(items){ lis = col.querySelectorAll('li'); items.forEach(function(t,i){ if(lis[i]) lis[i].textContent = t; }); }
    });
    card.querySelectorAll('.raw-data').forEach(function(rd){
      var h4 = rd.querySelector('h4'); if(h4) h4.textContent = T.raw_data;
      var p = rd.querySelector('p'); if(p) p.textContent = fd.raw;
    });
  });

  var cred = document.querySelector('.credit'); if(cred) cred.textContent = T.credit;
}
</script>

</body>
</html>"""

PATTERNS_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Instruction Patterns Guide — KoboToolbox XLSForm Converter</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  background: linear-gradient(135deg, #f0f4ff 0%, #e8f0fe 100%); min-height: 100vh;
  padding: 2rem 1rem; overflow-x: hidden; }
.container { background: white; border-radius: 16px; padding: 2.5rem; width: 100%;
  max-width: 1200px; margin: 0 auto; box-shadow: 0 4px 24px rgba(0,0,0,0.08); }
.back-link { display: inline-block; margin-bottom: 1.5rem; color: #4472C4;
  text-decoration: none; font-size: 0.85rem; font-weight: 600; }
.back-link:hover { text-decoration: underline; }
h1 { font-size: 1.7rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0.3rem; }
.subtitle { color: #6b7280; font-size: 0.9rem; margin-bottom: 1.75rem; }
.tabs { display: flex; gap: 0.5rem; flex-wrap: wrap; margin-bottom: 1.5rem;
  border-bottom: 2px solid #e5e7eb; padding-bottom: 0.5rem; }
.tab { padding: 0.5rem 1rem; border: 1px solid #e5e7eb; border-radius: 8px; cursor: pointer;
  font-size: 0.8rem; font-weight: 600; color: #6b7280; background: white; transition: all 0.15s; }
.tab:hover { border-color: #4472C4; color: #4472C4; background: #f0f4ff; }
.tab.active { background: #4472C4; color: white; border-color: #4472C4; }
.search-box { width: 100%; padding: 0.65rem 1rem; border: 1px solid #e5e7eb; border-radius: 8px;
  font-size: 0.88rem; margin-bottom: 1.25rem; outline: none; font-family: inherit; }
.search-box:focus { border-color: #4472C4; box-shadow: 0 0 0 3px rgba(68,114,196,0.1); }
table { width: 100%; border-collapse: collapse; font-size: 0.82rem; }
th { background: #1a3a6b; color: white; padding: 0.7rem 0.75rem; text-align: left;
  font-size: 0.78rem; position: sticky; top: 0; z-index: 1; }
td { padding: 0.6rem 0.75rem; border-bottom: 1px solid #f3f4f6; vertical-align: top; }
tr:nth-child(even) td { background: #f9fafb; }
tr:hover td { background: #f0f4ff; }
.num { color: #4472C4; font-weight: 700; text-align: center; white-space: nowrap; }
.pattern-name { font-weight: 600; color: #1a1a2e; }
.triggers { color: #374151; font-size: 0.77rem; }
.trigger-pill { display: inline-block; background: #f0f4ff; color: #4472C4; border: 1px solid #c7d2fe;
  border-radius: 4px; padding: 1px 6px; margin: 2px; font-size: 0.72rem; white-space: nowrap; }
.syntax { font-family: "Courier New", monospace; background: #f0f4ff; color: #1a1a2e;
  padding: 2px 6px; border-radius: 4px; font-size: 0.78rem; white-space: nowrap; }
.col-badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 0.7rem; font-weight: 700; }
.col-relevant { background: #d4edda; color: #155724; }
.col-constraint { background: #f8d7da; color: #721c24; }
.col-choice_filter { background: #fff3cd; color: #7d5a00; }
.col-calculation { background: #d1ecf1; color: #0c5460; }
.note-text { color: #6b7280; font-size: 0.75rem; font-style: italic; }
.info-box { background: #f8faff; border: 1px solid #e0e7ff; border-radius: 10px;
  padding: 1.25rem; margin-bottom: 1.5rem; font-size: 0.84rem; }
.info-box h3 { color: #1a3a6b; font-size: 0.9rem; margin-bottom: 0.75rem; }
.info-box ul { padding-left: 1.25rem; }
.info-box ul li { margin-bottom: 0.35rem; color: #374151; line-height: 1.6; }
.info-box code { background: #e8f0fe; padding: 1px 5px; border-radius: 3px;
  font-family: "Courier New", monospace; font-size: 0.8rem; color: #1a3a6b; }
.count { font-size: 0.78rem; color: #6b7280; margin-bottom: 0.75rem; }
.credit { text-align: center; margin-top: 2rem; font-size: 0.78rem; color: #9ca3af; font-style: italic; }
@media (max-width: 768px) { .container { padding: 1.25rem; } th, td { padding: 0.45rem 0.5rem; }
  .syntax { white-space: normal; } h1 { font-size: 1.3rem; } }
</style>
</head>
<body>
<div class="container">
  <a href="/" class="back-link">← Back to Converter</a>
  <h1>Instruction Patterns Guide</h1>
  <p class="subtitle">All metadata instructions the KoboToolbox XLSForm Converter can automatically detect and convert to valid XLSForm syntax.</p>

  <div class="info-box">
    <h3>📝 How to Write Instructions</h3>
    <ul>
      <li>Instructions appear as paragraphs <strong>above</strong> the question table: <code>Metadata type: Instruction</code></li>
      <li>All trigger words are <strong>case-insensitive</strong></li>
      <li>Leading zeros handled automatically — <code>01</code> and <code>1</code> produce the same result</li>
      <li>Prefixes <code>Que.</code> <code>Ques.</code> <code>Q.</code> <code>Question</code> before question names are stripped automatically</li>
      <li>Use <code>For [Q],</code> or <code>For [Q1] and [Q2],</code> to target a specific question in a funnel group</li>
      <li>Open-ended grid types: <code>Open ended grid</code> (text), <code>Integer grid</code> (integer), <code>Decimal grid</code> (decimal), <code>Date grid</code> (date)</li>
      <li>Sum constraint on grids: <code>Constraint: Sum of all topics must not exceed 10</code></li>
      <li>Max selections: <code>Constraint: Only select 5 options</code> → <code>count-selected(.) &lt;= 5</code></li>
      <li>Notes without a code are automatically named <code>NOTEA1</code>, <code>NOTEA2</code> etc.</li>
      <li>Use <code>{}</code> in a question label to display the label text of an option selected elsewhere</li>
    </ul>
  </div>

  <div class="tabs">
    <button class="tab active" onclick="filterCol('all',this)">All Patterns (52)</button>
    <button class="tab" onclick="filterCol('relevant',this)">Relevant (22)</button>
    <button class="tab" onclick="filterCol('constraint',this)">Constraint (17)</button>
    <button class="tab" onclick="filterCol('choice_filter',this)">Choice Filter (11)</button>
    <button class="tab" onclick="filterCol('special',this)">Special (2)</button>
  </div>

  <input class="search-box" type="text" placeholder="Search patterns, trigger words or syntax..."
    oninput="searchPatterns(this.value)" />
  <div class="count" id="countLabel">Showing all 52 patterns</div>

  <div style="overflow-x:auto;">
  <table id="patternsTable">
    <thead><tr>
      <th style="width:45px">#</th>
      <th style="width:190px">Pattern</th>
      <th>Trigger Words</th>
      <th style="width:240px">XLSForm Output</th>
      <th style="width:105px">Column</th>
      <th style="width:155px">Notes</th>
    </tr></thead>
    <tbody id="patternsBody"></tbody>
  </table>
  </div>
  <div class="credit">KoboToolbox XLSForm Converter — By Marvis Onyenwenu Enubiaka</div>
</div>

<script>
const patterns = [
  {num:"1",   name:"[Q] = [val]",                        triggers:["Display if","Ask if","Show if","Only show if","Show only if","Afficher si","Montrer si","Affiche si"],                                                                                                                          syntax:"${[Q]} = '[val]'",                                              col:"relevant",      notes:"Leading zeros removed"},
  {num:"2",   name:"[Q] != [val]",                        triggers:["Display if","Ask if","Skip if","Show if","Only show if","Show only if","Afficher si","Ignorer si","Passer si","Sauter si"],                                                                                                    syntax:"${[Q]} != '[val]'",                                             col:"relevant",      notes:""},
  {num:"3",   name:"Skip if [Q] = [val]",                 triggers:["Skip if","Ignorer si","Passer si","Sauter si"],                                                                                                                                                                                syntax:"${[Q]} != '[val]'",                                             col:"relevant",      notes:"Skip reverses the condition"},
  {num:"4",   name:"[Q1]=[val] or [Q2]=[val]",            triggers:["Display if","Ask if","Show if","Afficher si","Montrer si"],                                                                                                                                                                    syntax:"${[Q1]}='[v1]' or ${[Q2]}='[v2]'",                             col:"relevant",      notes:"Two different questions"},
  {num:"5",   name:"[Q1]=[val] and [Q2]=[val]",           triggers:["Display if","Ask if","Show if","Afficher si","Montrer si"],                                                                                                                                                                    syntax:"${[Q1]}='[v1]' and ${[Q2]}='[v2]'",                            col:"relevant",      notes:"Both must be true"},
  {num:"6",   name:"[Q]=[v1] or [v2] or [v3]",           triggers:["Display if","Ask if","Show if","Afficher si","Montrer si"],                                                                                                                                                                    syntax:"${[Q]}='v1' or ${[Q]}='v2' or ...",                            col:"relevant",      notes:"Multiple values same question"},
  {num:"7",   name:"[Q] is equal to [val]",               triggers:["Show if","Ask if","Only show if","Show only if","Montrer si","Demander si"],                                                                                                                                                   syntax:"${[Q]} = '[val]'",                                             col:"relevant",      notes:""},
  {num:"8",   name:"[Q] is not equal to [val]",           triggers:["Show if","Ask if","Skip if","Only show if","Montrer si","Demander si","Ignorer si"],                                                                                                                                           syntax:"${[Q]} != '[val]'",                                            col:"relevant",      notes:""},
  {num:"9",   name:"[Q] is [val]",                        triggers:["Display if","Ask if","Only show if","Show only if","Afficher si","Demander si","Montrer si"],                                                                                                                                  syntax:"${[Q]} = '[val]'",                                             col:"relevant",      notes:""},
  {num:"10",  name:"[Q] is not [val]",                    triggers:["Display if","Ask if","Only show if","Show only if","Afficher si","Demander si","Montrer si"],                                                                                                                                  syntax:"${[Q]} != '[val]'",                                            col:"relevant",      notes:""},
  {num:"11",  name:"[Q] is not = [val]",                  triggers:["Show if","Display if","Ask if","Skip if","Montrer si","Afficher si","Demander si","Ignorer si"],                                                                                                                               syntax:"${[Q]} != '[val]'",                                            col:"relevant",      notes:""},
  {num:"12",  name:"[Q] is = [val]",                      triggers:["Show if","Ask if","Only show if","Show only if","Montrer si","Demander si"],                                                                                                                                                   syntax:"${[Q]} = '[val]'",                                             col:"relevant",      notes:""},
  {num:"13",  name:"[val] is selected at [Q]",            triggers:["Display if","Ask if","Only show if","Show only if","Afficher si","Demander si","Montrer si"],                                                                                                                                  syntax:"selected(${[Q]}, '[val]')",                                    col:"relevant",      notes:""},
  {num:"14",  name:"[val] NOT selected at [Q]",           triggers:["Display if","Ask if","Only show if","Show only if","Afficher si","Demander si","Montrer si"],                                                                                                                                  syntax:"!selected(${[Q]}, '[val]')",                                   col:"relevant",      notes:""},
  {num:"15",  name:"[v1] or [v2] selected at [Q]",        triggers:["Display if","Ask if","Only show if","Afficher si","Demander si"],                                                                                                                                                             syntax:"selected(${[Q]},'v1') or selected(${[Q]},'v2')",               col:"relevant",      notes:""},
  {num:"16",  name:"[v1] and [v2] selected at [Q]",       triggers:["Display if","Ask if","Only show if","Afficher si","Demander si"],                                                                                                                                                             syntax:"selected(${[Q]},'v1') and selected(${[Q]},'v2')",              col:"relevant",      notes:"Both must be selected"},
  {num:"17",  name:"Display [Q] if [Q1]=[v] or [Q2]=[v]",triggers:["Display [Q] if","Afficher si","Demander si","Montrer si"],                                                                                                                                                                    syntax:"${[Q1]}='[val]' or ${[Q2]}='[val]'",                           col:"relevant",      notes:"Named Q only"},
  {num:"18",  name:"[Q] is answered / not empty",         triggers:["Display if","Ask if","Only show if","Show only if","Afficher si","Demander si","Montrer si"],                                                                                                                                  syntax:"${[Q]} != ''",                                                 col:"relevant",      notes:""},
  {num:"19",  name:"Topics for options at [Q]",           triggers:["Display topics for","For all... display topics","Afficher les sujets pour","Montrer les sujets pour"],                                                                                                                        syntax:"selected(${[Q]}, name)",                                       col:"relevant",      notes:"Grid topic relevant"},
  {num:"20",  name:"Ques. [list] if [val] at [Q]",        triggers:["Display Ques."],                                                                                                                                                                                                             syntax:"selected(${[Q]}, '[val]')",                                    col:"relevant",      notes:"Named questions only"},
  {num:"21",  name:"[Q1],[Q2],[Q3] if [val] at [Q]",      triggers:["Display [Q1],[Q2],[Q3] and [Qn] if [val] is selected at","Afficher si","Montrer si"],                                                                                                                                       syntax:"selected(${[Q]}, '[val]')",                                    col:"relevant",      notes:"Each named target"},
  {num:"22",  name:"Mixed equals + selected at",          triggers:["Display if [Q]=v1 or v2 and [val] is selected at [Q2]","Afficher si","Montrer si"],                                                                                                                                          syntax:"(${[Q]}='v1' or ${[Q]}='v2') and selected(${[Q2]},'[val]')",  col:"relevant",      notes:"Parentheses added automatically"},
  {num:"23",  name:"Do not proceed if [Q] = [val]",       triggers:["Do not proceed if","Ne pas continuer si","Ne continuez pas si"],                                                                                                                                                              syntax:"${[Q]} != '[val]'",                                            col:"constraint",    notes:"Must NOT equal val"},
  {num:"24",  name:"Do not proceed if [Q] != [val]",      triggers:["Do not proceed if","Ne pas continuer si","Ne continuez pas si"],                                                                                                                                                              syntax:"${[Q]} = '[val]'",                                             col:"constraint",    notes:"Must equal val"},
  {num:"25",  name:"Do not proceed if [Q] is [val]",      triggers:["Do not proceed if","Ne pas continuer si","Ne continuez pas si"],                                                                                                                                                              syntax:"${[Q]} != '[val]'",                                            col:"constraint",    notes:""},
  {num:"26",  name:"Do not proceed if [Q] is not = [val]",triggers:["Do not proceed if","Ne pas continuer si"],                                                                                                                                                                                   syntax:"${[Q]} = '[val]'",                                             col:"constraint",    notes:""},
  {num:"27",  name:"[val] selected at [Q]",               triggers:["Do not proceed if","Skip if","Ne pas continuer si"],                                                                                                                                                                         syntax:"!selected(${[Q]}, '[val]')",                                   col:"constraint",    notes:""},
  {num:"28",  name:"number > [val]",                      triggers:["Do not proceed if number >","Ne pas continuer si le nombre >"],                                                                                                                                                               syntax:". <= [val]",                                                   col:"constraint",    notes:""},
  {num:"29",  name:"number < [val]",                      triggers:["Do not proceed if number <","Ne pas continuer si le nombre <"],                                                                                                                                                               syntax:". >= [val]",                                                   col:"constraint",    notes:""},
  {num:"30",  name:"number >= [val]",                     triggers:["Do not proceed if number >=","Autoriser seulement un nombre entre","Permettre seulement un nombre entre"],                                                                                                                    syntax:". < [val]",                                                    col:"constraint",    notes:""},
  {num:"31",  name:"number <= [val]",                     triggers:["Do not proceed if number <=","Ne pas continuer si le nombre <="],                                                                                                                                                             syntax:". > [val]",                                                    col:"constraint",    notes:""},
  {num:"32",  name:"Number between [v1] and [v2]",        triggers:["Only allow a number between","Autoriser seulement un nombre entre","Permettre seulement un nombre entre"],                                                                                                                    syntax:". >= [v1] and . <= [v2]",                                      col:"constraint",    notes:""},
  {num:"33",  name:"[Q] = [Q2] variable comparison",      triggers:["Do not proceed if","Rendre [code] exclusif"],                                                                                                                                                                                syntax:"${[Q]} != ${[Q2]}",                                            col:"constraint",    notes:"Compares two questions"},
  {num:"34",  name:"Make [code] not selectable",          triggers:["Make [code] not selectable","For [Q] make [code] not selectable","Rendre [code] non sélectionnable","Rendre [code] impossible à sélectionner"],                                                                              syntax:". != '[code]'",                                                col:"constraint",    notes:"Any integer code"},
  {num:"35",  name:"Make [code] Exclusive",               triggers:["Make [code] Exclusive","Rendre [code] exclusif"],                                                                                                                                                                            syntax:"exclusive choice logic",                                       col:"constraint",    notes:"Deselects all others"},
  {num:"36",  name:"Only select [n] options",             triggers:["Only select [n] options","Select no more than [n] options","Do not select more than [n] options","Select at most [n] options","Maximum [n] selections","Maximum [n] options","Sélectionner seulement [n] options","Ne sélectionner que [n] options","Ne pas sélectionner plus de [n] options","Sélectionner au maximum [n] options","Sélectionner au plus [n] options"],  syntax:"count-selected(.) <= [n]",  col:"constraint",    notes:"select_multiple only"},
  {num:"37",  name:"Sum of topics must not exceed [n]",   triggers:["Sum of all topics must not exceed [n]","La somme de tous les sujets ne doit pas dépasser [n]"],                                                                                                                              syntax:"coalesce(sum) <= [n] on every topic",                          col:"constraint",    notes:"Grid questions"},
  {num:"38",  name:"Sum of topics must be = [n]",         triggers:["Sum of all topics must be = [n]","must equal [n]","must be equal to [n]","La somme de tous les sujets doit être = [n]","La somme de tous les sujets doit égaler [n]"],                                                       syntax:"sum<=n always; sum=n on last answered topic",                  col:"constraint",    notes:"Grid questions"},
  {num:"39",  name:"Sum of topics must be at least [n]",  triggers:["Sum of all topics must be at least [n]","La somme de tous les sujets doit être au moins [n]"],                                                                                                                               syntax:"coalesce(sum) >= [n] on every topic",                          col:"constraint",    notes:"Grid questions"},
  {num:"40",  name:"Options selected at [Q]",             triggers:["Display options selected at","Show options selected at","Only display options selected at","Display only options selected at","Filter to options selected at","Limit to options selected at","Show only options selected at","Show options chosen at","Afficher les options sélectionnées à","Montrer les options sélectionnées à","Afficher seulement les options sélectionnées à","Filtrer aux options sélectionnées à","Limiter aux options sélectionnées à","Montrer seulement les options sélectionnées à"], syntax:"selected(${[Q]}, name)", col:"choice_filter", notes:""},
  {num:"41",  name:"For [target], options at [source]",   triggers:["For [Q], display options selected at","Pour [Q], afficher les options sélectionnées à"],                                                                                                                                     syntax:"selected(${[source]}, name)",                                  col:"choice_filter", notes:"Named target only"},
  {num:"42",  name:"Options at [Q] for [Q1],[Q2]...",     triggers:["Display option(s) selected at [Q] for [Q1],[Q2]","Afficher les options sélectionnées à [Q] pour [Q1],[Q2]"],                                                                                                                 syntax:"selected(${[Q]}, name)",                                       col:"choice_filter", notes:"Same filter all targets"},
  {num:"43",  name:"Options at [Q1] or [Q2]",             triggers:["Display options selected at [Q1] or [Q2]","Afficher les options sélectionnées à [Q1] ou [Q2]"],                                                                                                                              syntax:"selected(${[Q1]}, name) or selected(${[Q2]}, name)",           col:"choice_filter", notes:"OR always union"},
  {num:"44",  name:"Options at [Q1] and [Q2] (same labels)",triggers:["Display options selected at [Q1] and [Q2]","Afficher les options sélectionnées à [Q1] et [Q2] (mêmes libellés)"],                                                                                                         syntax:"selected(${[Q1]}, name) and selected(${[Q2]}, name)",          col:"choice_filter", notes:"AND same labels=intersection"},
  {num:"45",  name:"Options at [Q1] and [Q2] (diff labels)",triggers:["Display options selected at [Q1] and [Q2]","Afficher les options sélectionnées à [Q1] et [Q2] (libellés différents)"],                                                                                                    syntax:"selected(${[Q1]}, name) or selected(${[Q2]}, name)",           col:"choice_filter", notes:"AND diff labels=union"},
  {num:"46",  name:"Options NOT selected at [Q]",         triggers:["Display options not selected at","Afficher les options non sélectionnées à"],                                                                                                                                                 syntax:"not(selected(${[Q]}, name))",                                  col:"choice_filter", notes:""},
  {num:"47",  name:"NOT selected at [Q1] and [Q2]",       triggers:["Display options not selected at [Q1] and [Q2]","Afficher les options non sélectionnées à [Q1] et [Q2]"],                                                                                                                    syntax:"not(selected(${[Q1]},name)) and not(selected(${[Q2]},name))", col:"choice_filter", notes:""},
  {num:"48",  name:"Always show option [code]",           triggers:["Make [code] always visible","Always show [code]","Always display [code]","Force display [code]","[code] should always appear","Toujours afficher [code]","Rendre [code] toujours visible","Toujours montrer [code]"],         syntax:"or name = '[code]'",                                           col:"choice_filter", notes:"Appended to filter"},
  {num:"49",  name:"Filter through question funnel",      triggers:["filter through question funnel","filter through funnel","Filtrer par entonnoir de questions","Filtrer à travers l'entonnoir de questions"],                                                                                    syntax:"Q2→selected(${Q1},name) Q3→selected(${Q2},name)...",          col:"choice_filter", notes:"Q1 never gets filter"},
  {num:"50",  name:"Link to selected [val] at [Q]",       triggers:["Link to all selected [val] at [Q]"],                                                                                                                                                                                         syntax:"selected-at(${[Q]}, position())",                              col:"calculation",   notes:""},
  {num:"51",  name:"Phone number regex",                  triggers:["write a regex syntax for [Country] phone number format","write regex for [Country] phone number","regex for [Country] phone number format","Phone number = [Country]","écrire une syntaxe regex pour le format de numéro de téléphone [Pays]","Numéro de téléphone = [Pays]"],  syntax:"regex(., '^...country-specific-regex...$')",  col:"constraint",    notes:"92 countries — <a href='/phone-countries' target='_blank' style='color:#4472C4'>View lookup →</a>; unknown → [REVIEW]"},
  {num:"52",  name:"Digit length constraint",             triggers:["Allow only [n] digits","Allow only [n] to [m] digits","Autoriser seulement [n] chiffres","Autoriser seulement [n] à [m] chiffres","N'autoriser que [n] chiffres","Permettre seulement [n] chiffres"],                         syntax:"regex(., '^[0-9]{n}$') or regex(., '^[0-9]{n,m}$')",          col:"constraint",    notes:"text & integer; exact or range; message auto-set"},
];

function colBadge(col){
  const l={relevant:'Relevant',constraint:'Constraint',choice_filter:'Choice Filter',calculation:'Calculation'};
  return `<span class="col-badge col-${col}">${l[col]||col}</span>`;
}
function renderTable(data){
  const tbody=document.getElementById('patternsBody');
  tbody.innerHTML='';
  data.forEach(p=>{
    const pills=p.triggers.map(t=>`<span class="trigger-pill">${t}</span>`).join(' ');
    tbody.innerHTML+=`<tr data-col="${p.col}">
      <td class="num">${p.num}</td>
      <td><div class="pattern-name">${p.name}</div></td>
      <td class="triggers">${pills}</td>
      <td><code class="syntax">${p.syntax}</code></td>
      <td>${colBadge(p.col)}</td>
      <td class="note-text">${p.notes}</td>
    </tr>`;
  });
  document.getElementById('countLabel').textContent=`Showing ${data.length} of ${patterns.length} patterns`;
}
let currentCol='all',currentSearch='';
function filterCol(col,btn){
  currentCol=col;
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  if(btn)btn.classList.add('active');
  applyFilters();
}
function searchPatterns(q){currentSearch=q.toLowerCase();applyFilters();}
function applyFilters(){
  let data=patterns;
  if(currentCol!=='all'&&currentCol!=='special') data=data.filter(p=>p.col===currentCol);
  else if(currentCol==='special') data=data.filter(p=>p.num==='44'||p.num==='45');
  if(currentSearch) data=data.filter(p=>
    p.name.toLowerCase().includes(currentSearch)||
    p.triggers.some(t=>t.toLowerCase().includes(currentSearch))||
    p.syntax.toLowerCase().includes(currentSearch)||
    p.notes.toLowerCase().includes(currentSearch)||
    p.col.toLowerCase().includes(currentSearch)
  );
  renderTable(data);
}

function applyPageLang(lang) {
  if (lang !== 'fr') return;
  var T = {
    title:    "Guide des modèles d'instructions",
    subtitle: "Toutes les instructions de métadonnées que le convertisseur peut détecter et convertir automatiquement en syntaxe XLSForm valide.",
    back:     "← Retour au convertisseur",
    tab_all:  "Tous les modèles",
    tab_rel:  "Pertinence",
    tab_con:  "Contrainte",
    tab_cf:   "Filtre de choix",
    tab_sp:   "Spéciaux",
    search_ph:"Rechercher des modèles, déclencheurs ou syntaxe...",
    th_num:   "#",
    th_name:  "Modèle",
    th_trig:  "Déclencheurs",
    th_out:   "Sortie XLSForm",
    th_col:   "Colonne",
    th_notes: "Notes",
    h3_how:   "📝 Comment écrire les instructions",
    info: [
      "Les instructions apparaissent comme des paragraphes <strong>au-dessus</strong> du tableau de question : <code>Type de métadonnée : Instruction</code>",
      "Tous les déclencheurs sont <strong>insensibles à la casse</strong>",
      "Les zéros initiaux sont gérés automatiquement — <code>01</code> et <code>1</code> donnent le même résultat",
      "Les préfixes <code>Que.</code> <code>Ques.</code> <code>Q.</code> <code>Question</code> avant les noms de questions sont supprimés automatiquement",
      "Utilisez <code>Pour [Q],</code> ou <code>Pour [Q1] et [Q2],</code> pour cibler une question spécifique dans un groupe en entonnoir",
      "Types de grille ouverte : <code>Open ended grid</code> (texte), <code>Integer grid</code> (entier), <code>Decimal grid</code> (décimal), <code>Date grid</code> (date)",
      "Contrainte de somme sur les grilles : <code>Constraint: Sum of all topics must not exceed 10</code>",
      "Sélections max : <code>Constraint: Only select 5 options</code> → <code>count-selected(.) &lt;= 5</code>",
      "Les notes sans code sont automatiquement nommées <code>NOTEA1</code>, <code>NOTEA2</code> etc.",
      "Utilisez <code>{}</code> dans un libellé de question pour afficher le texte d'une option sélectionnée ailleurs"
    ]
  };
  // Page title and subtitle
  var h1 = document.querySelector('h1'); if(h1) h1.textContent = T.title;
  var sub = document.querySelector('.subtitle'); if(sub) sub.textContent = T.subtitle;
  var bl = document.querySelector('.back-link'); if(bl) bl.textContent = T.back;
  // Search box placeholder
  var sb = document.querySelector('.search-box'); if(sb) sb.placeholder = T.search_ph;
  // Info box heading
  var ih = document.querySelector('.info-box h3'); if(ih) ih.textContent = T.h3_how;
  // Info box bullets
  var lis = document.querySelectorAll('.info-box ul li');
  T.info.forEach(function(txt, i) { if(lis[i]) lis[i].innerHTML = txt; });
  // Table headers
  var ths = document.querySelectorAll('th');
  var hdr = [T.th_num, T.th_name, T.th_trig, T.th_out, T.th_col, T.th_notes];
  ths.forEach(function(th,i){ if(hdr[i]) th.textContent = hdr[i]; });
  // Tab labels
  var tabs = document.querySelectorAll('.tab');
  var tlabels = [T.tab_all + ' (52)', T.tab_rel + ' (22)', T.tab_con + ' (17)', T.tab_cf + ' (11)', T.tab_sp + ' (2)'];
  tabs.forEach(function(t,i){ if(tlabels[i]) t.textContent = tlabels[i]; });
  // Count label
  var cl = document.getElementById('countLabel');
  if(cl) cl.textContent = 'Affichage de ' + patterns.length + ' modèles sur ' + patterns.length;
  // Override count label after filter
  var origRender = window._origRender || renderTable;
  window._origRender = origRender;
  window.renderTable = function(data) {
    origRender(data);
    var cl2 = document.getElementById('countLabel');
    if(cl2) cl2.textContent = 'Affichage de ' + data.length + ' modèles sur ' + patterns.length;
  };
}

renderTable(patterns);

</script>
</body>
</html>"""

PHONE_COUNTRIES_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Phone Number Country Lookup — KoboToolbox XLSForm Converter</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  background: linear-gradient(135deg, #f0f4ff 0%, #e8f0fe 100%);
  min-height: 100vh; padding: 2rem 1rem; overflow-x: hidden;
}
.container {
  background: white; border-radius: 16px; padding: 2.5rem;
  width: 100%; max-width: 1100px; margin: 0 auto;
  box-shadow: 0 4px 24px rgba(0,0,0,0.08);
}
.back-link {
  display: inline-block; margin-bottom: 1.5rem; color: #4472C4;
  text-decoration: none; font-size: 0.85rem; font-weight: 600;
}
.back-link:hover { text-decoration: underline; }
h1 { font-size: 1.7rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0.3rem; }
.subtitle { color: #6b7280; font-size: 0.9rem; margin-bottom: 1.5rem; }

/* ── Info boxes ── */
.info-grid {
  display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-bottom: 1.5rem;
}
.info-box {
  background: #f8faff; border: 1px solid #e0e7ff; border-radius: 10px;
  padding: 1.1rem 1.25rem; font-size: 0.84rem;
}
.info-box h3 { color: #1a3a6b; font-size: 0.88rem; margin-bottom: 0.6rem; font-weight: 700; }
.info-box p, .info-box li { color: #374151; line-height: 1.65; margin-bottom: 0.3rem; }
.info-box ul { padding-left: 1.1rem; }
.info-box code {
  background: #e8f0fe; padding: 1px 5px; border-radius: 3px;
  font-family: "Courier New", monospace; font-size: 0.79rem; color: #1a3a6b;
}
.digit-box {
  background: #f0fdf4; border: 1px solid #86efac; border-radius: 10px;
  padding: 1.1rem 1.25rem; font-size: 0.84rem; margin-bottom: 1.5rem;
}
.digit-box h3 { color: #166534; font-size: 0.88rem; margin-bottom: 0.6rem; font-weight: 700; }
.digit-box p { color: #374151; line-height: 1.65; margin-bottom: 0.5rem; }
.digit-box .eg-row {
  display: flex; align-items: flex-start; gap: 0.75rem; margin-bottom: 0.4rem;
  font-size: 0.82rem;
}
.digit-box .eg-input {
  background: #dcfce7; border: 1px solid #86efac; border-radius: 5px;
  padding: 2px 8px; font-family: "Courier New", monospace; color: #166534;
  white-space: nowrap; flex-shrink: 0;
}
.digit-box .eg-arrow { color: #6b7280; flex-shrink: 0; padding-top: 2px; }
.digit-box .eg-output {
  font-family: "Courier New", monospace; color: #374151; font-size: 0.78rem;
  word-break: break-all;
}

/* ── Controls ── */
.controls {
  display: flex; gap: 0.75rem; align-items: center;
  flex-wrap: wrap; margin-bottom: 1rem;
}
.search-box {
  flex: 1; min-width: 180px; padding: 0.6rem 1rem;
  border: 1px solid #e5e7eb; border-radius: 8px;
  font-size: 0.88rem; outline: none; font-family: inherit;
}
.search-box:focus { border-color: #4472C4; box-shadow: 0 0 0 3px rgba(68,114,196,0.1); }
.region-select {
  padding: 0.6rem 0.9rem; border: 1px solid #e5e7eb; border-radius: 8px;
  font-size: 0.84rem; outline: none; font-family: inherit; cursor: pointer;
  background: white; color: #374151;
}
.region-select:focus { border-color: #4472C4; }
.count-label { font-size: 0.78rem; color: #6b7280; white-space: nowrap; }

/* ── Table ── */
.table-wrap { overflow-x: auto; }
table { width: 100%; border-collapse: collapse; font-size: 0.82rem; }
th {
  background: #1a3a6b; color: white; padding: 0.65rem 0.75rem;
  text-align: left; font-size: 0.78rem; position: sticky; top: 0; z-index: 1;
  white-space: nowrap;
}
td { padding: 0.55rem 0.75rem; border-bottom: 1px solid #f3f4f6; vertical-align: top; }
tr:nth-child(even) td { background: #f9fafb; }
tr:hover td { background: #f0f4ff; }
.col-country { font-weight: 600; color: #1a1a2e; white-space: nowrap; }
.col-trigger {
  font-family: "Courier New", monospace; font-size: 0.76rem; color: #374151;
}
.trigger-pill {
  display: inline-block; background: #f0f4ff; color: #4472C4;
  border: 1px solid #c7d2fe; border-radius: 4px; padding: 1px 6px;
  margin: 1px; font-size: 0.72rem; white-space: nowrap;
}
.col-regex {
  font-family: "Courier New", monospace; font-size: 0.72rem; color: #374151;
  word-break: break-all; max-width: 320px;
}
.col-msg { font-size: 0.78rem; color: #6b7280; white-space: nowrap; }
.copy-btn {
  background: none; border: 1px solid #e5e7eb; border-radius: 4px;
  padding: 1px 7px; font-size: 0.68rem; cursor: pointer; color: #6b7280;
  margin-left: 4px; font-family: inherit; transition: all 0.15s;
  white-space: nowrap;
}
.copy-btn:hover { background: #f0f4ff; border-color: #4472C4; color: #4472C4; }
.copy-btn.copied { background: #f0fdf4; border-color: #86efac; color: #166534; }
.no-results {
  text-align: center; padding: 2.5rem; color: #9ca3af; font-size: 0.88rem;
}
.credit {
  text-align: center; margin-top: 2rem; font-size: 0.78rem;
  color: #9ca3af; font-style: italic;
}
@media (max-width: 700px) {
  .info-grid { grid-template-columns: 1fr; }
  .container { padding: 1.25rem 1rem; }
  h1 { font-size: 1.3rem; }
}
</style>
</head>
<body>
<div class="container">
  <a href="/" class="back-link">← Back to Converter</a>
  <h1>Phone Number Country Lookup</h1>
  <p class="subtitle">All 92 countries supported by Pattern 50. Use the trigger phrase exactly as shown to generate the correct regex constraint automatically.</p>

  <div class="info-grid">
    <div class="info-box">
      <h3>📞 How to use Pattern 50</h3>
      <ul>
        <li>Write as a <strong>Constraint</strong> paragraph above the question table</li>
        <li>Format: <code>Constraint: write a regex syntax for [Country] phone number format</code></li>
        <li>Use either the adjective <em>or</em> noun trigger — both work</li>
        <li>The constraint message is set automatically</li>
        <li>Country not listed? You'll get a <code>[REVIEW]</code> flag to add it manually</li>
      </ul>
    </div>
    <div class="info-box">
      <h3>🇫🇷 French trigger</h3>
      <ul>
        <li>Format: <code>Constraint: écrire une syntaxe regex pour le format de numéro de téléphone [Pays]</code></li>
        <li>Use the French adjective from the trigger column</li>
        <li>The constraint message in French: <em>Doit être au format de numéro de téléphone [Pays]</em></li>
        <li>Works identically to the English trigger</li>
      </ul>
    </div>
  </div>

  <div class="digit-box">
    <h3>🔢 Pattern 51 — Digit Length Constraint (standalone)</h3>
    <p>Use this on any <strong>text</strong> or <strong>integer</strong> question to restrict the number of digits allowed. Works independently — no phone regex needed.</p>
    <div class="eg-row">
      <span class="eg-input">Constraint: Allow only 11 digits</span>
      <span class="eg-arrow">→</span>
      <span class="eg-output">regex(., '^[0-9]{11}$') &nbsp;|&nbsp; Must be exactly 11 digits</span>
    </div>
    <div class="eg-row">
      <span class="eg-input">Constraint: Allow only 10 to 13 digits</span>
      <span class="eg-arrow">→</span>
      <span class="eg-output">regex(., '^[0-9]{10,13}$') &nbsp;|&nbsp; Must be between 10 and 13 digits</span>
    </div>
    <div class="eg-row">
      <span class="eg-input" style="background:#d1fae5;border-color:#6ee7b7;">French: Autoriser seulement 11 chiffres</span>
      <span class="eg-arrow">→</span>
      <span class="eg-output">regex(., '^[0-9]{11}$') &nbsp;|&nbsp; Must be exactly 11 digits</span>
    </div>
    <div class="eg-row">
      <span class="eg-input" style="background:#d1fae5;border-color:#6ee7b7;">French: Autoriser seulement 10 à 13 chiffres</span>
      <span class="eg-arrow">→</span>
      <span class="eg-output">regex(., '^[0-9]{10,13}$') &nbsp;|&nbsp; Must be between 10 and 13 digits</span>
    </div>
  </div>

  <div class="controls">
    <input class="search-box" type="text" id="searchBox"
      placeholder="Search country name or trigger word..." oninput="applyFilters()" />
    <select class="region-select" id="regionSelect" onchange="applyFilters()">
      <option value="all">All regions</option>
      <option value="Africa">Africa</option>
      <option value="Middle East">Middle East</option>
      <option value="South Asia">South Asia</option>
      <option value="Southeast Asia">Southeast Asia</option>
      <option value="East Asia">East Asia</option>
      <option value="Europe">Europe</option>
      <option value="Americas">Americas</option>
      <option value="Oceania">Oceania</option>
    </select>
    <span class="count-label" id="countLabel">Showing all 92 countries</span>
  </div>

  <div class="table-wrap">
    <table id="countryTable">
      <thead><tr>
        <th style="width:140px">Country</th>
        <th style="width:220px">Trigger words</th>
        <th>Regex constraint</th>
        <th style="width:240px">Constraint message</th>
      </tr></thead>
      <tbody id="tableBody"></tbody>
    </table>
    <div class="no-results" id="noResults" style="display:none;">No countries match your search.</div>
  </div>

  <div class="credit">KoboToolbox XLSForm Converter — By Marvis Onyenwenu Enubiaka</div>
</div>

<script>
const COUNTRIES = [
  {name:"Algerian",primary:"alg\\u00e9rienne",secondary:"alg\\u00e9rien",regex:"^(\\\\+213|0)(5[0-9]|6[0-9]|7[0-9])[0-9]{7}$",msg:"Must be in Algerian phone number format"},
  {name:"American",primary:"united states",secondary:"am\\u00e9ricaine",regex:"^(\\\\+1)?[2-9][0-9]{2}[2-9][0-9]{6}$",msg:"Must be in American phone number format"},
  {name:"Angolan",primary:"angolan",secondary:"angola",regex:"^(\\\\+244)(9[1-9][0-9])[0-9]{6}$",msg:"Must be in Angolan phone number format"},
  {name:"Argentinian",primary:"argentinian",secondary:"argentina",regex:"^(\\\\+54|0)?(9)?[1-9][0-9]{9}$",msg:"Must be in Argentinian phone number format"},
  {name:"Australian",primary:"australienne",secondary:"australien",regex:"^(\\\\+61|0)(4[0-9]{2})[0-9]{6}$",msg:"Must be in Australian phone number format"},
  {name:"Bangladeshi",primary:"bangladeshi",secondary:"bangladesh",regex:"^(\\\\+880|0)(1[3-9][0-9])[0-9]{7}$",msg:"Must be in Bangladeshi phone number format"},
  {name:"Belgian",primary:"belgium",secondary:"belgian",regex:"^(\\\\+32|0)(4[5-9][0-9])[0-9]{6}$",msg:"Must be in Belgian phone number format"},
  {name:"Beninese",primary:"beninese",secondary:"benin",regex:"^(\\\\+229)(9[0-9]|6[0-9])[0-9]{6}$",msg:"Must be in Beninese phone number format"},
  {name:"Botswanan",primary:"botswanan",secondary:"botswana",regex:"^(\\\\+267)(7[0-9])[0-9]{6}$",msg:"Must be in Botswanan phone number format"},
  {name:"Brazilian",primary:"br\\u00e9silienne",secondary:"br\\u00e9silien",regex:"^(\\\\+55|0)?([1-9]{2})(9[1-9][0-9]{3}|[2-5][0-9]{3})[0-9]{4}$",msg:"Must be in Brazilian phone number format"},
  {name:"British",primary:"united kingdom",secondary:"britannique",regex:"^(\\\\+44|0)(7[0-9]{3})[0-9]{6}$",msg:"Must be in British phone number format"},
  {name:"Burkinab\\u00e9",primary:"burkina faso",secondary:"burkinabe",regex:"^(\\\\+226)(6[0-9]|7[0-9])[0-9]{6}$",msg:"Must be in Burkinab\\u00e9 phone number format"},
  {name:"Burundian",primary:"burundian",secondary:"burundi",regex:"^(\\\\+257)(7[0-9]|6[0-9])[0-9]{6}$",msg:"Must be in Burundian phone number format"},
  {name:"Cambodian",primary:"cambodian",secondary:"cambodia",regex:"^(\\\\+855|0)(1[0-9]|6[0-9]|7[0-9]|8[0-9]|9[0-9])[0-9]{6,7}$",msg:"Must be in Cambodian phone number format"},
  {name:"Cameroonian",primary:"camerounaise",secondary:"camerounais",regex:"^(\\\\+237)(6[5-9][0-9])[0-9]{6}$",msg:"Must be in Cameroonian phone number format"},
  {name:"Canadian",primary:"canadian",secondary:"canada",regex:"^(\\\\+1)?[2-9][0-9]{2}[2-9][0-9]{6}$",msg:"Must be in Canadian phone number format"},
  {name:"Cape Verdean",primary:"cape verdean",secondary:"cape verde",regex:"^(\\\\+238)(9[0-9]|5[0-9])[0-9]{5}$",msg:"Must be in Cape Verdean phone number format"},
  {name:"Central African",primary:"central african",secondary:"car",regex:"^(\\\\+236)(7[0-9])[0-9]{6}$",msg:"Must be in Central African phone number format"},
  {name:"Chadian",primary:"chadian",secondary:"chad",regex:"^(\\\\+235)(6[0-9]|9[0-9])[0-9]{6}$",msg:"Must be in Chadian phone number format"},
  {name:"Chinese",primary:"chinese",secondary:"china",regex:"^(\\\\+86|0)?(1[3-9][0-9])[0-9]{8}$",msg:"Must be in Chinese phone number format"},
  {name:"Colombian",primary:"colombian",secondary:"colombia",regex:"^(\\\\+57)?(3[0-9]{2})[0-9]{7}$",msg:"Must be in Colombian phone number format"},
  {name:"Congolese",primary:"democratic republic of congo",secondary:"congolaise",regex:"^(\\\\+243|0)(8[0-9]|9[0-9])[0-9]{7}$",msg:"Must be in Congolese phone number format"},
  {name:"Congolese (Republic)",primary:"republic of congo",secondary:"congo brazzaville",regex:"^(\\\\+242)(0[4-6][0-9])[0-9]{6}$",msg:"Must be in Congolese (Republic) phone number format"},
  {name:"Danish",primary:"denmark",secondary:"danish",regex:"^(\\\\+45)?[2-9][0-9]{7}$",msg:"Must be in Danish phone number format"},
  {name:"Djiboutian",primary:"djiboutian",secondary:"djibouti",regex:"^(\\\\+253)(7[7-9])[0-9]{6}$",msg:"Must be in Djiboutian phone number format"},
  {name:"Dutch",primary:"netherlands",secondary:"holland",regex:"^(\\\\+31|0)(6)[0-9]{8}$",msg:"Must be in Dutch phone number format"},
  {name:"Egyptian",primary:"\\u00e9gyptienne",secondary:"\\u00e9gyptien",regex:"^(\\\\+20|0)(10|11|12|15)[0-9]{8}$",msg:"Must be in Egyptian phone number format"},
  {name:"Emirati",primary:"united arab emirates",secondary:"emirati",regex:"^(\\\\+971|0)(5[0-9])[0-9]{7}$",msg:"Must be in Emirati phone number format"},
  {name:"Equatorial Guinean",primary:"equatorial guinean",secondary:"equatorial guinea",regex:"^(\\\\+240)(2[2-9]|5[5-9])[0-9]{7}$",msg:"Must be in Equatorial Guinean phone number format"},
  {name:"Eritrean",primary:"eritrean",secondary:"eritrea",regex:"^(\\\\+291|0)(7[0-9])[0-9]{5}$",msg:"Must be in Eritrean phone number format"},
  {name:"Ethiopian",primary:"\\u00e9thiopienne",secondary:"\\u00e9thiopien",regex:"^(\\\\+251|0)(9[0-9]|7[0-9])[0-9]{7}$",msg:"Must be in Ethiopian phone number format"},
  {name:"Filipino",primary:"philippines",secondary:"filipino",regex:"^(\\\\+63|0)(9[0-9]{2})[0-9]{7}$",msg:"Must be in Filipino phone number format"},
  {name:"French",primary:"fran\\u00e7aise",secondary:"fran\\u00e7ais",regex:"^(\\\\+33|0)(6[0-9]|7[0-9])[0-9]{7}$",msg:"Must be in French phone number format"},
  {name:"Gabonese",primary:"gabonese",secondary:"gabon",regex:"^(\\\\+241)(0[6-7][0-9])[0-9]{5}$",msg:"Must be in Gabonese phone number format"},
  {name:"Gambian",primary:"gambian",secondary:"gambia",regex:"^(\\\\+220)(7[0-9]|9[0-9]|3[0-9]|6[0-9])[0-9]{5}$",msg:"Must be in Gambian phone number format"},
  {name:"German",primary:"germany",secondary:"german",regex:"^(\\\\+49|0)(1[5-7][0-9])[0-9]{7,8}$",msg:"Must be in German phone number format"},
  {name:"Ghanaian",primary:"ghan\\u00e9enne",secondary:"ghanaian",regex:"^(\\\\+233|0)(20|23|24|25|26|27|28|50|54|55|57|59)[0-9]{7}$",msg:"Must be in Ghanaian phone number format"},
  {name:"Guinean",primary:"guinean",secondary:"guinea",regex:"^(\\\\+224)(6[0-9]{2})[0-9]{6}$",msg:"Must be in Guinean phone number format"},
  {name:"Indian",primary:"indienne",secondary:"indien",regex:"^(\\\\+91|0)?[6-9][0-9]{9}$",msg:"Must be in Indian phone number format"},
  {name:"Indonesian",primary:"indonesian",secondary:"indonesia",regex:"^(\\\\+62|0)(8[1-9][0-9])[0-9]{6,9}$",msg:"Must be in Indonesian phone number format"},
  {name:"Iranian",primary:"iranian",secondary:"iran",regex:"^(\\\\+98|0)(9[0-9]{2})[0-9]{7}$",msg:"Must be in Iranian phone number format"},
  {name:"Iraqi",primary:"iraqi",secondary:"iraq",regex:"^(\\\\+964|0)(7[3-9][0-9])[0-9]{7}$",msg:"Must be in Iraqi phone number format"},
  {name:"Italian",primary:"italian",secondary:"italy",regex:"^(\\\\+39|0)?(3[0-9]{2})[0-9]{7}$",msg:"Must be in Italian phone number format"},
  {name:"Ivorian",primary:"c\\u00f4te d'ivoire",secondary:"cote d'ivoire",regex:"^(\\\\+225|0)(07|05|01)[0-9]{8}$",msg:"Must be in Ivorian phone number format"},
  {name:"Japanese",primary:"japanese",secondary:"japan",regex:"^(\\\\+81|0)(70|80|90)[0-9]{8}$",msg:"Must be in Japanese phone number format"},
  {name:"Jordanian",primary:"jordanian",secondary:"jordan",regex:"^(\\\\+962|0)(7[5-9])[0-9]{7}$",msg:"Must be in Jordanian phone number format"},
  {name:"Kenyan",primary:"k\\u00e9nyane",secondary:"k\\u00e9nyan",regex:"^(\\\\+254|0)(7[0-9]|1[0-9])[0-9]{7}$",msg:"Must be in Kenyan phone number format"},
  {name:"Lesotho",primary:"lesotho",secondary:"",regex:"^(\\\\+266)(5[0-9]|6[0-9])[0-9]{6}$",msg:"Must be in Lesotho phone number format"},
  {name:"Liberian",primary:"liberian",secondary:"liberia",regex:"^(\\\\+231|0)(77|88|55)[0-9]{6}$",msg:"Must be in Liberian phone number format"},
  {name:"Libyan",primary:"libyan",secondary:"libya",regex:"^(\\\\+218|0)(9[1-5])[0-9]{7}$",msg:"Must be in Libyan phone number format"},
  {name:"Malagasy",primary:"madagascar",secondary:"malagasy",regex:"^(\\\\+261|0)(3[2-4])[0-9]{7}$",msg:"Must be in Malagasy phone number format"},
  {name:"Malawian",primary:"malawian",secondary:"malawi",regex:"^(\\\\+265|0)(9[0-9]|8[0-9])[0-9]{7}$",msg:"Must be in Malawian phone number format"},
  {name:"Malian",primary:"malienne",secondary:"malien",regex:"^(\\\\+223)(6[0-9]|7[0-9]|8[0-9]|9[0-9])[0-9]{6}$",msg:"Must be in Malian phone number format"},
  {name:"Mauritian",primary:"mauritius",secondary:"mauritian",regex:"^(\\\\+230)(5[0-9])[0-9]{6}$",msg:"Must be in Mauritian phone number format"},
  {name:"Mexican",primary:"mexican",secondary:"mexico",regex:"^(\\\\+52)?(1)?[2-9][0-9]{9}$",msg:"Must be in Mexican phone number format"},
  {name:"Moroccan",primary:"marocaine",secondary:"moroccan",regex:"^(\\\\+212|0)(6[0-9]|7[0-9])[0-9]{7}$",msg:"Must be in Moroccan phone number format"},
  {name:"Mozambican",primary:"mozambique",secondary:"mozambican",regex:"^(\\\\+258)(8[0-9]|7[0-9])[0-9]{7}$",msg:"Must be in Mozambican phone number format"},
  {name:"Myanmar",primary:"myanmar",secondary:"burmese",regex:"^(\\\\+95|0)(9[0-9]{1,2})[0-9]{6,8}$",msg:"Must be in Myanmar phone number format"},
  {name:"Namibian",primary:"namibian",secondary:"namibia",regex:"^(\\\\+264|0)(8[0-9]|6[0-9])[0-9]{7}$",msg:"Must be in Namibian phone number format"},
  {name:"Nepali",primary:"nepali",secondary:"nepal",regex:"^(\\\\+977|0)(9[78][0-9])[0-9]{7}$",msg:"Must be in Nepali phone number format"},
  {name:"New Zealand",primary:"new zealander",secondary:"new zealand",regex:"^(\\\\+64|0)(2[0-9])[0-9]{6,8}$",msg:"Must be in New Zealand phone number format"},
  {name:"Nigerian",primary:"nig\\u00e9riane",secondary:"nig\\u00e9rian",regex:"^(\\\\+234|0)(70|80|90|81|91)[0-9]{8}$",msg:"Must be in Nigerian phone number format"},
  {name:"Nigerien",primary:"nigerien",secondary:"niger",regex:"^(\\\\+227)(9[0-9]|8[0-9])[0-9]{6}$",msg:"Must be in Nigerien phone number format"},
  {name:"Norwegian",primary:"norwegian",secondary:"norway",regex:"^(\\\\+47)?(4[0-9]|9[0-9])[0-9]{6}$",msg:"Must be in Norwegian phone number format"},
  {name:"Pakistani",primary:"pakistani",secondary:"pakistan",regex:"^(\\\\+92|0)(3[0-9]{2})[0-9]{7}$",msg:"Must be in Pakistani phone number format"},
  {name:"Polish",primary:"polish",secondary:"poland",regex:"^(\\\\+48)?(5[0-9]|6[0-9]|7[0-9]|8[0-9])[0-9]{7}$",msg:"Must be in Polish phone number format"},
  {name:"Portuguese",primary:"portuguese",secondary:"portugal",regex:"^(\\\\+351)?(9[1-6])[0-9]{7}$",msg:"Must be in Portuguese phone number format"},
  {name:"Russian",primary:"russian",secondary:"russia",regex:"^(\\\\+7|8)(9[0-9]{2})[0-9]{7}$",msg:"Must be in Russian phone number format"},
  {name:"Rwandan",primary:"rwandaise",secondary:"rwandais",regex:"^(\\\\+250|0)(7[2-9])[0-9]{7}$",msg:"Must be in Rwandan phone number format"},
  {name:"Saudi",primary:"saudi arabian",secondary:"saudi arabia",regex:"^(\\\\+966|0)(5[0-9])[0-9]{7}$",msg:"Must be in Saudi phone number format"},
  {name:"Senegalese",primary:"s\\u00e9n\\u00e9galaise",secondary:"s\\u00e9n\\u00e9galais",regex:"^(\\\\+221)(7[0-9]|6[0-9])[0-9]{7}$",msg:"Must be in Senegalese phone number format"},
  {name:"Seychellois",primary:"seychellois",secondary:"seychelles",regex:"^(\\\\+248)(2[5-9]|5[0-9])[0-9]{5}$",msg:"Must be in Seychellois phone number format"},
  {name:"Sierra Leonean",primary:"sierra leonean",secondary:"sierra leone",regex:"^(\\\\+232|0)(7[0-9]|8[0-9]|9[0-9]|3[0-9])[0-9]{6}$",msg:"Must be in Sierra Leonean phone number format"},
  {name:"Somali",primary:"somalia",secondary:"somali",regex:"^(\\\\+252|0)(6[0-9]|7[0-9]|9[0-9])[0-9]{6}$",msg:"Must be in Somali phone number format"},
  {name:"South African",primary:"sud-africaine",secondary:"south african",regex:"^(\\\\+27|0)(6[0-9]|7[0-9]|8[0-9])[0-9]{7}$",msg:"Must be in South African phone number format"},
  {name:"South Korean",primary:"south korean",secondary:"south korea",regex:"^(\\\\+82|0)(10|11)[0-9]{7,8}$",msg:"Must be in South Korean phone number format"},
  {name:"South Sudanese",primary:"south sudanese",secondary:"south sudan",regex:"^(\\\\+211|0)(9[0-9])[0-9]{7}$",msg:"Must be in South Sudanese phone number format"},
  {name:"Spanish",primary:"spanish",secondary:"spain",regex:"^(\\\\+34)?(6[0-9]{2}|7[0-9]{2})[0-9]{6}$",msg:"Must be in Spanish phone number format"},
  {name:"Sri Lankan",primary:"sri lankan",secondary:"sri lanka",regex:"^(\\\\+94|0)(7[0-9])[0-9]{7}$",msg:"Must be in Sri Lankan phone number format"},
  {name:"Sudanese",primary:"sudanese",secondary:"sudan",regex:"^(\\\\+249|0)(9[0-9])[0-9]{7}$",msg:"Must be in Sudanese phone number format"},
  {name:"Swazi",primary:"swaziland",secondary:"eswatini",regex:"^(\\\\+268)(7[6-9]|6[0-9])[0-9]{6}$",msg:"Must be in Swazi phone number format"},
  {name:"Swedish",primary:"swedish",secondary:"sweden",regex:"^(\\\\+46|0)(7[0-9])[0-9]{7}$",msg:"Must be in Swedish phone number format"},
  {name:"Tanzanian",primary:"tanzanienne",secondary:"tanzanien",regex:"^(\\\\+255|0)(7[0-9]|6[0-9])[0-9]{7}$",msg:"Must be in Tanzanian phone number format"},
  {name:"Thai",primary:"thailand",secondary:"thai",regex:"^(\\\\+66|0)(6[0-9]|8[0-9]|9[0-9])[0-9]{7}$",msg:"Must be in Thai phone number format"},
  {name:"Togolese",primary:"togolese",secondary:"togo",regex:"^(\\\\+228)(9[0-9]|7[0-9])[0-9]{6}$",msg:"Must be in Togolese phone number format"},
  {name:"Tunisian",primary:"tunisienne",secondary:"tunisien",regex:"^(\\\\+216)(2[0-9]|5[0-9]|9[0-9])[0-9]{6}$",msg:"Must be in Tunisian phone number format"},
  {name:"Turkish",primary:"t\\u00fcrkiye",secondary:"turkish",regex:"^(\\\\+90|0)(5[0-9]{2})[0-9]{7}$",msg:"Must be in Turkish phone number format"},
  {name:"Ugandan",primary:"ougandaise",secondary:"ougandais",regex:"^(\\\\+256|0)(7[0-9]|4[0-9])[0-9]{7}$",msg:"Must be in Ugandan phone number format"},
  {name:"Ukrainian",primary:"ukrainian",secondary:"ukraine",regex:"^(\\\\+380|0)(6[0-9]|7[0-9]|9[0-9])[0-9]{7}$",msg:"Must be in Ukrainian phone number format"},
  {name:"Vietnamese",primary:"vietnamese",secondary:"vietnam",regex:"^(\\\\+84|0)(3[2-9]|5[6-9]|7[0-9]|8[1-9]|9[0-9])[0-9]{7}$",msg:"Must be in Vietnamese phone number format"},
  {name:"Zambian",primary:"zambian",secondary:"zambia",regex:"^(\\\\+260|0)(9[5-9]|7[6-9])[0-9]{7}$",msg:"Must be in Zambian phone number format"},
  {name:"Zimbabwean",primary:"zimbabwean",secondary:"zimbabwe",regex:"^(\\\\+263|0)(7[1-8])[0-9]{7}$",msg:"Must be in Zimbabwean phone number format"}
];

// Region map — assign each display_name to a region
const REGION_MAP = {
  // Africa
  "Nigerian":"Africa","Ghanaian":"Africa","Senegalese":"Africa","Ivorian":"Africa",
  "Cameroonian":"Africa","Malian":"Africa","Burkinabé":"Africa","Guinean":"Africa",
  "Sierra Leonean":"Africa","Liberian":"Africa","Gambian":"Africa","Togolese":"Africa",
  "Beninese":"Africa","Nigerien":"Africa","Cape Verdean":"Africa",
  "Kenyan":"Africa","Tanzanian":"Africa","Ugandan":"Africa","Rwandan":"Africa",
  "Ethiopian":"Africa","Burundian":"Africa","Mozambican":"Africa","Malagasy":"Africa",
  "Malawian":"Africa","Zambian":"Africa","Zimbabwean":"Africa",
  "South African":"Africa","Namibian":"Africa","Botswanan":"Africa","Lesotho":"Africa",
  "Swazi":"Africa","Mauritian":"Africa","Seychellois":"Africa",
  "Egyptian":"Africa","Moroccan":"Africa","Tunisian":"Africa","Algerian":"Africa",
  "Libyan":"Africa","Sudanese":"Africa",
  "Congolese":"Africa","Congolese (Republic)":"Africa","Angolan":"Africa",
  "Chadian":"Africa","Central African":"Africa","Gabonese":"Africa",
  "Equatorial Guinean":"Africa","Somali":"Africa","Djiboutian":"Africa",
  "Eritrean":"Africa","South Sudanese":"Africa",
  // Middle East
  "Saudi":"Middle East","Emirati":"Middle East","Jordanian":"Middle East",
  "Iraqi":"Middle East","Iranian":"Middle East",
  // South Asia
  "Indian":"South Asia","Pakistani":"South Asia","Bangladeshi":"South Asia",
  "Sri Lankan":"South Asia","Nepali":"South Asia",
  // Southeast Asia
  "Indonesian":"Southeast Asia","Filipino":"Southeast Asia","Vietnamese":"Southeast Asia",
  "Thai":"Southeast Asia","Myanmar":"Southeast Asia","Cambodian":"Southeast Asia",
  // East Asia
  "Chinese":"East Asia","Japanese":"East Asia","South Korean":"East Asia",
  // Europe
  "British":"Europe","French":"Europe","German":"Europe","Italian":"Europe",
  "Spanish":"Europe","Portuguese":"Europe","Dutch":"Europe","Belgian":"Europe",
  "Swedish":"Europe","Norwegian":"Europe","Danish":"Europe","Polish":"Europe",
  "Turkish":"Europe","Ukrainian":"Europe","Russian":"Europe",
  // Americas
  "American":"Americas","Canadian":"Americas","Mexican":"Americas",
  "Brazilian":"Americas","Colombian":"Americas","Argentinian":"Americas",
  // Oceania
  "Australian":"Oceania","New Zealand":"Oceania",
};

function copyText(text, btn) {
  navigator.clipboard.writeText(text).then(() => {
    btn.textContent = '✓ Copied';
    btn.classList.add('copied');
    setTimeout(() => { btn.textContent = 'Copy'; btn.classList.remove('copied'); }, 1500);
  });
}

function renderTable(data) {
  const tbody = document.getElementById('tableBody');
  const noRes = document.getElementById('noResults');
  tbody.innerHTML = '';
  if (data.length === 0) {
    noRes.style.display = 'block';
    document.getElementById('countLabel').textContent = 'No results';
    return;
  }
  noRes.style.display = 'none';
  data.forEach(row => {
    const pills = [row.primary, row.secondary].filter(Boolean)
      .map(t => `<span class="trigger-pill">${t}</span>`).join(' ');
    const regexDisplay = row.regex.replace(/</g,'&lt;').replace(/>/g,'&gt;');
    tbody.innerHTML += `<tr>
      <td class="col-country">${row.name}</td>
      <td class="col-trigger">${pills}</td>
      <td class="col-regex">
        <span id="rx_${row.name.replace(/[^a-z]/gi,'_')}">${regexDisplay}</span>
        <button class="copy-btn" onclick="copyText(${JSON.stringify(row.regex)}, this)">Copy</button>
      </td>
      <td class="col-msg">${row.msg}</td>
    </tr>`;
  });
  document.getElementById('countLabel').textContent =
    `Showing ${data.length} of ${COUNTRIES.length} countries`;
}

let currentSearch = '', currentRegion = 'all';

function applyFilters() {
  currentSearch = document.getElementById('searchBox').value.toLowerCase().trim();
  currentRegion = document.getElementById('regionSelect').value;
  let data = COUNTRIES;
  if (currentRegion !== 'all') {
    data = data.filter(r => REGION_MAP[r.name] === currentRegion);
  }
  if (currentSearch) {
    data = data.filter(r =>
      r.name.toLowerCase().includes(currentSearch) ||
      r.primary.toLowerCase().includes(currentSearch) ||
      (r.secondary && r.secondary.toLowerCase().includes(currentSearch)) ||
      r.regex.toLowerCase().includes(currentSearch)
    );
  }
  renderTable(data);
}

renderTable(COUNTRIES);

function applyPageLang(lang) {
  if (lang !== 'fr') return;
  var T = {
    title:    "Recherche de pays — numéros de téléphone",
    subtitle: "Les 92 pays pris en charge par le modèle 50. Utilisez la phrase déclencheur exactement comme indiqué pour générer automatiquement la bonne contrainte regex.",
    back:     "← Retour au convertisseur",
    h3_how:   "📞 Comment utiliser le modèle 50",
    info_en:  [
      "Écrivez comme paragraphe de <strong>Contrainte</strong> au-dessus du tableau de question",
      "Format : <code>Constraint: write a regex syntax for [Pays] phone number format</code>",
      "Utilisez le déclencheur adjectif <em>ou</em> nominal — les deux fonctionnent",
      "Le message de contrainte est défini automatiquement",
      "Pays non listé ? Vous recevrez un marqueur <code>[REVIEW]</code> pour l'ajouter manuellement"
    ],
    h3_fr:    "🇫🇷 Déclencheur français",
    info_fr:  [
      "Format : <code>Constraint: écrire une syntaxe regex pour le format de numéro de téléphone [Pays]</code>",
      "Ou plus simplement : <code>Constraint: Numéro de téléphone = [Pays]</code>",
      "Utilisez l'adjectif français de la colonne déclencheur",
      "Le message de contrainte en français : <em>Doit être au format de numéro de téléphone [Pays]</em>",
      "Fonctionne de manière identique au déclencheur anglais"
    ],
    h3_digit: "🔢 Modèle 51 — Contrainte de longueur en chiffres (autonome)",
    digit_p:  "Utilisez sur toute question <strong>texte</strong> ou <strong>entier</strong> pour restreindre le nombre de chiffres autorisés. Fonctionne indépendamment — aucune regex de numéro de téléphone nécessaire.",
    search_ph:"Rechercher un pays, un déclencheur ou un code...",
    regions:  ["Toutes les régions","Afrique","Moyen-Orient","Asie du Sud","Asie du Sud-Est","Asie de l'Est","Europe","Amériques","Océanie"],
    th:       ["Pays","Déclencheurs","Contrainte regex","Message de contrainte"],
    copy_btn: "Copier",
    copied:   "✓ Copié",
    no_res:   "Aucun pays ne correspond à votre recherche.",
    credit:   "Convertisseur XLSForm KoboToolbox — Par Marvis Onyenwenu Enubiaka",
    showing:  function(n,t){ return "Affichage de " + n + " sur " + t + " pays"; },
    showing_all: function(t){ return "Affichage de tous les " + t + " pays"; }
  };

  var h1 = document.querySelector('h1'); if(h1) h1.textContent = T.title;
  var sub = document.querySelector('.subtitle'); if(sub) sub.textContent = T.subtitle;
  var bl = document.querySelector('.back-link'); if(bl) bl.textContent = T.back;

  // Info boxes headings and bullets
  var infoBoxes = document.querySelectorAll('.info-box');
  if(infoBoxes[0]){
    var h3 = infoBoxes[0].querySelector('h3'); if(h3) h3.textContent = T.h3_how;
    var lis = infoBoxes[0].querySelectorAll('li');
    T.info_en.forEach(function(t,i){ if(lis[i]) lis[i].innerHTML = t; });
  }
  if(infoBoxes[1]){
    var h3 = infoBoxes[1].querySelector('h3'); if(h3) h3.textContent = T.h3_fr;
    var lis = infoBoxes[1].querySelectorAll('li');
    T.info_fr.forEach(function(t,i){ if(lis[i]) lis[i].innerHTML = t; });
  }

  // Digit box
  var dbox = document.querySelector('.digit-box');
  if(dbox){
    var dh3 = dbox.querySelector('h3'); if(dh3) dh3.textContent = T.h3_digit;
    var dp = dbox.querySelector('p'); if(dp) dp.innerHTML = T.digit_p;
  }

  // Search box placeholder
  var sb = document.getElementById('searchBox'); if(sb) sb.placeholder = T.search_ph;

  // Region select options
  var rs = document.getElementById('regionSelect');
  if(rs){ rs.querySelectorAll('option').forEach(function(o,i){ if(T.regions[i]) o.textContent = T.regions[i]; }); }

  // Table headers
  var ths = document.querySelectorAll('th');
  T.th.forEach(function(t,i){ if(ths[i]) ths[i].textContent = t; });

  // Count label
  var cl = document.getElementById('countLabel');
  if(cl) cl.textContent = T.showing_all(COUNTRIES.length);

  // Copy button text (patch renderTable to translate copy buttons)
  var origRT = window.renderTable;
  window.renderTable = function(data) {
    origRT(data);
    document.querySelectorAll('.copy-btn').forEach(function(btn){
      if(!btn.classList.contains('copied')) btn.textContent = T.copy_btn;
    });
    var cl2 = document.getElementById('countLabel');
    if(cl2){
      cl2.textContent = data.length === COUNTRIES.length
        ? T.showing_all(COUNTRIES.length)
        : T.showing(data.length, COUNTRIES.length);
    }
  };

  // No results text
  var nr = document.getElementById('noResults'); if(nr) nr.textContent = T.no_res;

  // Credit
  var cred = document.querySelector('.credit'); if(cred) cred.textContent = T.credit;
}

</script>
</body>
</html>"""


HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KoboToolbox XLSForm Converter</title>
    <!-- Google Analytics -->
    <script async src="https://www.googletagmanager.com/gtag/js?id=G-W6NBFZRNFG"></script>
    <script>
      window.dataLayer = window.dataLayer || [];
      function gtag(){dataLayer.push(arguments);}
      gtag('js', new Date());
      gtag('config', 'G-W6NBFZRNFG');
    </script>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: linear-gradient(135deg, #f0f4ff 0%, #e8f0fe 100%);
            min-height: 100vh;
            display: flex;
            align-items: flex-start;
            justify-content: center;
            padding: 2rem 1rem;
            overflow-x: hidden;
        }
        .container {
            background: white;
            border-radius: 16px;
            padding: 2.5rem;
            width: 100%;
            max-width: 860px;
            box-shadow: 0 4px 24px rgba(0,0,0,0.08);
            overflow: hidden;
        }
        h1 { font-size: 1.8rem; font-weight: 700; color: #1a1a2e; text-align: center; margin-bottom: 0.4rem; }
        .subtitle { text-align: center; color: #6b7280; font-size: 0.9rem; margin-bottom: 2rem; }
        .main-layout {
            display: flex;
            gap: 2rem;
            align-items: flex-start;
            margin-bottom: 1.5rem;
        }
        .left-panel { flex: 1; min-width: 0; }
        .right-panel { flex-shrink: 0; width: auto; }

        /* ── Responsive — tablet ────────────────────────────── */
        @media (max-width: 700px) {
            body { padding: 1rem 0.75rem; }
            .container { padding: 1.5rem 1.25rem; border-radius: 12px; }
            h1 { font-size: 1.4rem; }
            .subtitle { font-size: 0.82rem; margin-bottom: 1.25rem; }
            .main-layout {
                flex-direction: column;
                gap: 1.25rem;
            }
            .left-panel { width: 100%; }
            .right-panel { width: 100%; }
            .repeat-options { width: 100%; }
            .repeat-option { width: 100%; }
        }

        /* ── Responsive — mobile ────────────────────────────── */
        @media (max-width: 480px) {
            body { padding: 0.75rem 0.5rem; }
            .container { padding: 1.25rem 1rem; border-radius: 10px; }
            h1 { font-size: 1.25rem; }
            .upload-area { padding: 1.5rem 1rem; }
            .upload-area h3 { font-size: 0.9rem; }
            .convert-btn {
                font-size: 0.95rem;
                padding: 0.85rem;
            }
            .file-name-row { font-size: 0.8rem; }
            .output-options { padding: 0.75rem 1rem; }
        }
        .upload-area {
            border: 2px dashed #c7d2fe;
            border-radius: 12px;
            padding: 2rem 1.5rem;
            text-align: center;
            cursor: pointer;
            transition: all 0.2s ease;
            background: #fafbff;
            position: relative;
            margin-bottom: 0.75rem;
        }
        .upload-area:hover, .upload-area.dragover { border-color: #4472C4; background: #f0f4ff; }
        .upload-icon { font-size: 2.5rem; margin-bottom: 0.5rem; }
        .upload-area h3 { font-size: 1rem; color: #374151; margin-bottom: 0.25rem; }
        .upload-area p { font-size: 0.82rem; color: #9ca3af; }
        .file-input { position: absolute; inset: 0; width: 100%; height: 100%; opacity: 0; cursor: pointer; }
        .selected-file {
            display: flex; align-items: center; gap: 0.5rem;
            padding: 0.5rem 0.75rem; background: #f3f4f6;
            border-radius: 8px; font-size: 0.82rem; color: #374151;
            margin-bottom: 0.75rem; min-height: 36px;
        }
        .convert-btn {
            width: 100%; padding: 0.85rem; background: #4472C4;
            color: white; border: none; border-radius: 10px;
            font-size: 1rem; font-weight: 600; cursor: pointer;
            transition: all 0.2s ease;
        }
        .convert-btn:hover:not(:disabled) { background: #3461b0; }
        .convert-btn:disabled { background: #9ca3af; cursor: not-allowed; }
        .repeat-section { width: 100%; }
        .repeat-section label.section-label {
            display: block; font-size: 0.82rem; font-weight: 600;
            color: #374151; margin-bottom: 0.4rem;
        }
        .repeat-options { display: flex; flex-direction: column; gap: 0.3rem; width: fit-content; }
        .repeat-option {
            display: flex; align-items: center; gap: 0.55rem;
            padding: 0.5rem 0.75rem; border: 1px solid #e5e7eb;
            border-radius: 7px; cursor: pointer;
            transition: all 0.15s ease; position: relative;
            white-space: nowrap;
        }
        .repeat-option:hover { border-color: #4472C4; background: #f8faff; }
        .repeat-option input[type="radio"] { accent-color: #4472C4; flex-shrink: 0; margin: 0; }
        .repeat-option .option-title { font-size: 0.8rem; font-weight: 600; color: #1f2937; }
        .badge-default {
            font-size: 0.62rem; background: #4472C4; color: white;
            padding: 1px 5px; border-radius: 10px; font-weight: 600;
            margin-left: 4px; vertical-align: middle;
        }
        .repeat-option:has(input:checked) { border-color: #4472C4; background: #f0f4ff; }
        .repeat-option:has(input:checked) .option-title { color: #4472C4; }
        .repeat-option .tooltip-text {
            visibility: hidden; opacity: 0; width: 240px;
            white-space: normal; word-wrap: break-word;
            background: #1f2937; color: #fff; font-size: 0.72rem;
            line-height: 1.5; border-radius: 6px; padding: 8px 11px;
            position: absolute; z-index: 200;
            bottom: calc(100% + 10px); left: 0;
            transition: opacity 0.2s; pointer-events: none;
            box-shadow: 0 4px 14px rgba(0,0,0,0.25);
        }
        .repeat-option .tooltip-text::after {
            content: ""; position: absolute; top: 100%; left: 20px;
            border: 6px solid transparent; border-top-color: #1f2937;
        }
        .repeat-option:hover .tooltip-text { visibility: visible; opacity: 1; }
        .output-options {
            margin-top: 1.2rem; padding: 0.9rem 1.1rem;
            background: #f8faff; border: 1px solid #e5e7eb; border-radius: 10px;
        }
        .output-options .section-label {
            display: block; font-size: 0.82rem; font-weight: 600;
            color: #374151; margin-bottom: 0.6rem;
        }
        .opt-row {
            display: inline-flex; align-items: center; gap: 0.55rem;
            padding: 0.45rem 0.65rem; border: 1px solid #e5e7eb;
            border-radius: 7px; cursor: pointer; position: relative;
            transition: all 0.15s ease; white-space: nowrap;
        }
        .opt-row:hover { border-color: #4472C4; background: #f0f4ff; }
        .opt-row input[type="checkbox"] { accent-color: #4472C4; flex-shrink: 0; }
        .opt-row .opt-label { font-size: 0.8rem; font-weight: 600; color: #1f2937; }
        .opt-row:has(input:checked) { border-color: #4472C4; background: #f0f4ff; }
        .opt-row .opt-tooltip {
            visibility: hidden; opacity: 0; width: 260px;
            white-space: normal; word-wrap: break-word;
            background: #1f2937; color: #fff; font-size: 0.72rem;
            line-height: 1.5; border-radius: 6px; padding: 8px 11px;
            position: absolute; z-index: 200;
            bottom: calc(100% + 10px); left: 0;
            transition: opacity 0.2s; pointer-events: none;
            box-shadow: 0 4px 14px rgba(0,0,0,0.25);
        }
        .opt-row .opt-tooltip::after {
            content: ""; position: absolute; top: 100%; left: 20px;
            border: 6px solid transparent; border-top-color: #1f2937;
        }
        .opt-row:hover .opt-tooltip { visibility: visible; opacity: 1; }
        .status-box {
            margin-top: 1.2rem; padding: 1rem 1.25rem; border-radius: 10px;
            font-size: 0.88rem; line-height: 1.6; display: none; text-align: center;
        }
        .status-box.loading { background: #fefce8; border: 1px solid #fde68a; color: #854d0e; display: block; }
        .status-box.success { background: #f0fdf4; border: 1px solid #86efac; color: #166534; display: block; }
        .status-box.error   { background: #fef2f2; border: 1px solid #fca5a5; color: #991b1b; display: block; }
        .download-btn {
            display: inline-block; margin-top: 0.75rem; padding: 0.6rem 1.5rem;
            background: #4472C4; color: white; border-radius: 8px;
            text-decoration: none; font-weight: 600; font-size: 0.88rem;
        }
        .download-btn:hover { background: #3461b0; }
        .spinner {
            display: inline-block; width: 14px; height: 14px;
            border: 2px solid #fde68a; border-top-color: #854d0e;
            border-radius: 50%; animation: spin 0.7s linear infinite;
            margin-right: 6px; vertical-align: middle;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        .rules-note {
            margin-top: 1.5rem; padding-top: 1rem; border-top: 1px solid #f3f4f6;
            font-size: 0.8rem; color: #9ca3af; text-align: center; line-height: 1.6;
        }
        .credit {
            margin-top: 0.75rem; padding-top: 0.75rem; border-top: 1px solid #f3f4f6;
            text-align: center; font-size: 0.75rem; color: #c4c4c4;
            font-style: italic; letter-spacing: 0.03em;
        }
    </style>
</head>
<body>
<div class="container">

    <h1>KoboToolbox XLSForm Converter</h1>
    <p class="subtitle">Upload a Word questionnaire and get a ready-to-use XLSForm</p>

    <div class="main-layout">

        <div class="left-panel">
            <div class="upload-area" id="uploadArea">
                <div class="upload-icon">📄</div>
                <h3>Drop your Word document here</h3>
                <p><span data-i18n="drop_or">or click to browse</span> <span data-i18n="drop_formats">— .docx, .doc, .pdf</span></p>
                <input type="file" id="fileInput" class="file-input" accept=".docx,.doc,.pdf" />
            </div>
            <div class="selected-file" id="selectedFile">
                <span>📎</span>
                <span id="fileName">No file selected</span>
            </div>
            <div style="display:flex; justify-content:center; margin-bottom:0.75rem;">
                <div style="display:inline-flex; border:1px solid #e5e7eb; border-radius:24px; overflow:hidden; background:#f9fafb;">
                    <button id="langBtnEn" onclick="setLanguage('en')"
                        style="padding:0.4rem 1.1rem; font-size:0.82rem; font-weight:600;
                               border:none; cursor:pointer; font-family:inherit;
                               background:#4472C4; color:white; transition:all 0.15s;">
                        🇬🇧 English
                    </button>
                    <button id="langBtnFr" onclick="setLanguage('fr')"
                        style="padding:0.4rem 1.1rem; font-size:0.82rem; font-weight:600;
                               border:none; cursor:pointer; font-family:inherit;
                               background:transparent; color:#6b7280; transition:all 0.15s;">
                        🇫🇷 Français
                    </button>
                </div>
            </div>
            <button class="convert-btn" id="convertBtn" disabled>
                Convert to XLSForm
            </button>
        </div>

        <div class="right-panel">
            <div class="repeat-section">
                <label class="section-label">🔁 Repeat Group Format</label>
                <div class="repeat-options">

                    <label class="repeat-option">
                        <input type="radio" name="repeatFormat" value="4" checked />
                        <div class="option-title" data-i18n="smart">🧠 Smart Auto-Select <span class="badge-default" data-i18n="default_badge">Default</span></div>
                        <span class="tooltip-text">The engine evaluates each repeat group individually — considering the number of options, how the study is structured, and what works best for the raw data output. Different groups in the same questionnaire may use different formats.</span>
                    </label>

                    <label class="repeat-option">
                        <input type="radio" name="repeatFormat" value="3" />
                        <div class="option-title" data-i18n="format3">Automated Repeat Loop</div>
                        <span class="tooltip-text">True begin_repeat driven by a prior select_multiple. Adapts automatically per respondent. Best for large or variable lists.</span>
                    </label>

                    <label class="repeat-option">
                        <input type="radio" name="repeatFormat" value="2" />
                        <div class="option-title" data-i18n="format2">Brand-Fixed Group Loop</div>
                        <span class="tooltip-text">Each brand gets its own permanently fixed group that appears conditionally when selected. Best for brand equity tracking, retail audits and competitive benchmarking.</span>
                    </label>

                    <label class="repeat-option">
                        <input type="radio" name="repeatFormat" value="1" />
                        <div class="option-title" data-i18n="format1">Sequential Positional Loop</div>
                        <span class="tooltip-text">Fixed unrolled slot groups — enumerator picks one brand per slot. Already-selected brands are excluded from subsequent slots. Best for product testing, sensory evaluation and ranked preference studies.</span>
                    </label>

                    <label class="repeat-option">
                        <input type="radio" name="repeatFormat" value="5" />
                        <div class="option-title" data-i18n="direct">✍️ Direct Scripting Mode</div>
                        <span class="tooltip-text">No system repeat groups are applied. Questions are scripted exactly as written in the questionnaire — group by group, question by question, with no automated repeat wrapping.</span>
                    </label>

                </div>
            </div>
        </div>

    </div>

    <div class="status-box" id="statusBox"></div>

    <div style="text-align:center; margin: 0.5rem 0 1rem 0;">
        <a href="/repeat-guide" target="_blank"
           style="font-size:0.78rem; color:#4472C4; text-decoration:none; font-weight:500;">
            <span data-i18n="repeat_guide_inline">📖 Not sure which repeat format to use? View the Repeat Group Standards Guide →</span>
        </a>
    </div>

    <div class="output-options">
        <span class="section-label">⚙️ Output Options</span>
        <div style="display:flex; gap:0.75rem; flex-wrap:wrap;">
            <label class="opt-row">
                <input type="checkbox" id="includeCode" name="includeCode" checked />
                <span class="opt-label">Include question code in label</span>
                <span class="opt-tooltip">When checked, question labels include the question code — e.g. "RSD1a. What is your name?". When unchecked, only the question text appears — e.g. "What is your name?". Recommended to keep checked for easier data reference.</span>
            </label>
            <label class="opt-row">
                <input type="checkbox" id="boldHints" name="boldHints" checked />
                <span class="opt-label">Bold hint text</span>
                <span class="opt-tooltip">When checked, hint text is wrapped in bold markdown (**text**) so it displays larger and easier to read on KoboToolbox, ODK and SurveyCTO. Recommended for field surveys on mobile devices.</span>
            </label>
        </div>
    </div>

    <div class="rules-note">
        <span data-i18n="rules_note">Applies all KoboToolbox XLSForm rules automatically —
        sections, groups, grids, SO questions, choice filters and more.</span>
    </div>

    <div class="credit">By Marvis Onyenwenu Enubiaka</div>
    <div style="text-align:center; margin-top:0.75rem; font-size:0.8rem; display:flex; gap:0.75rem 1.5rem; justify-content:center; flex-wrap:wrap; padding:0 0.5rem;">
        <a href="/repeat-guide" target="_blank" style="color:#4472C4; text-decoration:none; font-weight:600;">
            <span data-i18n="repeat_link">📖 Repeat Group Standards Guide</span>
        </a>
        <a href="/patterns" target="_blank" style="color:#4472C4; text-decoration:none; font-weight:600;">
            <span data-i18n="patterns_link">📋 Instruction Patterns Guide</span>
        </a>
        <a href="/phone-countries" target="_blank" style="color:#4472C4; text-decoration:none; font-weight:600;">
            <span data-i18n="phone_link">📞 Phone Number Country Lookup</span>
        </a>
        <a href="#" id="youtubeLink" target="_blank" style="color:#e53e3e; text-decoration:none; font-weight:600;">
            <span data-i18n="youtube_link">▶ Learn the EQDS on YouTube</span>
        </a>
        <a href="javascript:void(0)" onclick="openReview()" style="color:#059669; text-decoration:none; font-weight:600;">
            <span data-i18n="rate_link">⭐ Rate This Tool</span>
        </a>
    </div>

    <!-- Review Modal -->
    <div id="reviewModal" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,0.5); z-index:1000; align-items:center; justify-content:center;">
        <div style="background:white; border-radius:16px; padding:2rem; width:100%; max-width:460px; margin:1rem;">
            <h3 style="font-size:1.1rem; color:#1a1a2e; margin-bottom:0.5rem;">Rate This Tool</h3>
            <p style="font-size:0.82rem; color:#6b7280; margin-bottom:1rem;">Your feedback helps improve the KoboToolbox XLSForm Converter.</p>
            <div id="starRow" style="font-size:2rem; margin-bottom:1rem; cursor:pointer; display:flex; gap:0.25rem;">
                <span class="star" data-val="1">☆</span>
                <span class="star" data-val="2">☆</span>
                <span class="star" data-val="3">☆</span>
                <span class="star" data-val="4">☆</span>
                <span class="star" data-val="5">☆</span>
            </div>
            <textarea id="reviewComment" placeholder="Write a comment (optional)..."
                style="width:100%; padding:0.65rem; border:1px solid #e5e7eb; border-radius:8px;
                font-size:0.85rem; font-family:inherit; resize:vertical; min-height:80px;
                margin-bottom:0.75rem; outline:none;"></textarea>
            <div style="display:flex; gap:0.75rem;">
                <button onclick="submitReview()"
                    style="flex:1; padding:0.7rem; background:#4472C4; color:white;
                    border:none; border-radius:8px; font-weight:600; cursor:pointer; font-size:0.9rem;">
                    Submit Review
                </button>
                <button onclick="closeReview()"
                    style="padding:0.7rem 1rem; background:#f3f4f6; color:#374151;
                    border:none; border-radius:8px; font-weight:600; cursor:pointer; font-size:0.9rem;">
                    Cancel
                </button>
            </div>
            <div id="reviewStatus" style="margin-top:0.75rem; font-size:0.82rem; text-align:center;"></div>
        </div>
    </div>

</div>

<script>
    // ── Language state (global — needed for onclick handlers) ────────────────
    let currentLang = localStorage.getItem('xlsconv_lang') || 'en';
    const UI = {
        en:{
            converting:'⏳ Converting your questionnaire... please wait.',
            success:'✅ Conversion successful! Your XLSForm is ready.',
            download:'⬇ Download XLSForm',
            warnings_pre:'⚠',
            warnings_post:'issue(s) found — please review:',
            error_pre:'❌ Error:',
            convert_btn:'Convert to XLSForm',
            title:'KoboToolbox XLSForm Converter',
            subtitle:'Upload a Word questionnaire and get a ready-to-use XLSForm',
            repeat_label:'🔁 Repeat Group Format',
            output_options:'⚙️ Output Options',
            include_code:'Include question code in label',
            bold_hints:'Bold hint text',
            no_file:'No file selected',
            drop_title:'Drop your Word document here',
            drop_or:'or click to browse',
            drop_formats:'— .docx, .doc, .pdf',
            repeat_guide_inline:'📖 Not sure which repeat format to use? View the Repeat Group Standards Guide →',
            smart:'🧠 Smart Auto-Select',
            default_badge:'Default',
            format3:'Automated Repeat Loop',
            format2:'Brand-Fixed Group Loop',
            format1:'Sequential Positional Loop',
            direct:'✍️ Direct Scripting Mode',
            rules_note:'Applies all KoboToolbox XLSForm rules automatically — sections, groups, grids, SO questions, choice filters and more.',
            repeat_link:'📖 Repeat Group Standards Guide',
            patterns_link:'📋 Instruction Patterns Guide',
            youtube_link:'▶ Learn the EQDS on YouTube',
            rate_link:'⭐ Rate This Tool',
            phone_link:'📞 Phone Number Country Lookup',
        },
        fr:{
            converting:'⏳ Conversion de votre questionnaire... veuillez patienter.',
            success:'✅ Conversion réussie ! Votre XLSForm est prêt.',
            download:'⬇ Télécharger le XLSForm',
            warnings_pre:'⚠',
            warnings_post:'problème(s) trouvé(s) — veuillez vérifier :',
            error_pre:'❌ Erreur :',
            convert_btn:'Convertir en XLSForm',
            title:'Convertisseur XLSForm KoboToolbox',
            subtitle:"Téléversez un questionnaire Word et obtenez un XLSForm prêt à l'emploi",
            repeat_label:'🔁 Format de groupe répété',
            output_options:'⚙️ Options de sortie',
            include_code:"Inclure le code de question dans l'étiquette",
            bold_hints:"Texte d'indice en gras",
            no_file:'Aucun fichier sélectionné',
            drop_title:'Déposez votre document Word ici',
            drop_or:'ou cliquez pour parcourir',
            drop_formats:'— .docx, .doc, .pdf',
            repeat_guide_inline:"📖 Vous ne savez pas quel format ? Voir le Guide des groupes répétés →",
            smart:'🧠 Sélection automatique intelligente',
            default_badge:'Défaut',
            format3:'Boucle répétée automatisée',
            format2:'Boucle de groupe fixe',
            format1:'Boucle positionnelle séquentielle',
            direct:'✍️ Mode de script direct',
            rules_note:"Applique automatiquement toutes les règles XLSForm de KoboToolbox — sections, groupes, grilles, questions SO, filtres de choix et plus.",
            repeat_link:"📖 Guide des normes de groupes répétés",
            patterns_link:"📋 Guide des modèles d'instructions",
            youtube_link:"▶ Apprendre l'EQDS sur YouTube",
            rate_link:'⭐ Évaluer cet outil',
            phone_link:'📞 Recherche de pays — numéros de téléphone',
        }
    };


    function setLanguage(lang){
        currentLang = lang;
        localStorage.setItem('xlsconv_lang', lang);
        const enBtn = document.getElementById('langBtnEn');
        const frBtn = document.getElementById('langBtnFr');
        if(enBtn){ enBtn.style.background = lang==='en' ? '#4472C4' : 'transparent'; enBtn.style.color = lang==='en' ? 'white' : '#6b7280'; }
        if(frBtn){ frBtn.style.background = lang==='fr' ? '#4472C4' : 'transparent'; frBtn.style.color = lang==='fr' ? 'white' : '#6b7280'; }
        applyLanguage(lang);
    }
    function applyLanguage(lang){
        const t = UI[lang];
        const h1 = document.querySelector('h1'); if(h1) h1.textContent = t.title;
        const sub = document.querySelector('.subtitle'); if(sub) sub.textContent = t.subtitle;
        const cb = document.getElementById('convertBtn'); if(cb) cb.textContent = t.convert_btn;
        const rl = document.querySelector('.repeat-section .section-label'); if(rl) rl.textContent = t.repeat_label;
        const ol = document.querySelector('.output-options .section-label'); if(ol) ol.textContent = t.output_options;
        const opts = document.querySelectorAll('.opt-row .opt-label');
        if(opts[0]) opts[0].textContent = t.include_code;
        if(opts[1]) opts[1].textContent = t.bold_hints;
        const dt = document.querySelector('.upload-area h3'); if(dt) dt.textContent = t.drop_title;
        const fn = document.getElementById('fileName');
        if(fn && (fn.textContent===UI.en.no_file||fn.textContent===UI.fr.no_file)) fn.textContent = t.no_file;
        // Translate all data-i18n elements
        document.querySelectorAll('[data-i18n]').forEach(el => {
            const key = el.getAttribute('data-i18n');
            if(t[key] !== undefined) el.textContent = t[key];
        });
        // Translate Rate This Tool modal
        var rateModal = {
            en: {
                title:   'Rate This Tool',
                body:    'Your feedback helps improve the KoboToolbox XLSForm Converter.',
                ph:      'Write a comment (optional)...',
                submit:  'Submit Review',
                cancel:  'Cancel'
            },
            fr: {
                title:   'Évaluer cet outil',
                body:    'Vos commentaires aident à améliorer le convertisseur XLSForm KoboToolbox.',
                ph:      'Écrire un commentaire (facultatif)...',
                submit:  "Soumettre l'avis",
                cancel:  'Annuler'
            }
        };
        var rm = rateModal[lang] || rateModal.en;
        var rh = document.querySelector('#reviewModal h3');  if(rh) rh.textContent = rm.title;
        var rp = document.querySelector('#reviewModal p');   if(rp) rp.textContent = rm.body;
        var rc = document.getElementById('reviewComment');   if(rc) rc.placeholder  = rm.ph;
        document.querySelectorAll('#reviewModal button').forEach(function(btn){
            var txt = btn.textContent.trim();
            if(txt === 'Submit Review' || txt === "Soumettre l'avis") btn.textContent = rm.submit;
            if(txt === 'Cancel'        || txt === 'Annuler')           btn.textContent = rm.cancel;
        });
        // Update footer links to include lang= parameter
        document.querySelectorAll('a[href^="/repeat-guide"], a[href^="/patterns"], a[href^="/phone-countries"]').forEach(function(a){
            var base = a.getAttribute('href').split('?')[0];
            a.setAttribute('href', lang === 'fr' ? base + '?lang=fr' : base);
        });
    }

    function setFile(file) {
        if (!file) return;
        const allowed = ['.docx', '.doc', '.pdf'];
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        if (!allowed.includes(ext)) {
            showStatus('error', '❌ Unsupported file type. Please upload a .docx, .doc, or .pdf file.');
            return;
        }
        selectedFile = file;
        fileName.textContent = file.name;
        fileName.style.color = '#374151';
        fileName.style.fontWeight = '600';
        convertBtn.disabled = false;
        statusBox.className = "status-box";
        statusBox.innerHTML = "";
    }

    // ── Wait for DOM to be fully ready ──────────────────────────────────────
    document.addEventListener("DOMContentLoaded", function() {

    // ── Restore saved language preference ────────────────────────────────────
    if (currentLang === 'fr') { setLanguage('fr'); }

    // ── File input change (click to browse) ─────────────────────────────────
    fileInput.addEventListener("change", function() {
        if (this.files && this.files[0]) {
            setFile(this.files[0]);
        }
    });

    // ── Drag and drop ────────────────────────────────────────────────────────
    uploadArea.addEventListener("dragover", function(e) {
        e.preventDefault();
        e.stopPropagation();
        uploadArea.classList.add("dragover");
    });
    uploadArea.addEventListener("dragleave", function(e) {
        e.preventDefault();
        e.stopPropagation();
        uploadArea.classList.remove("dragover");
    });
    uploadArea.addEventListener("drop", function(e) {
        e.preventDefault();
        e.stopPropagation();
        uploadArea.classList.remove("dragover");
        const files = e.dataTransfer.files;
        if (files && files.length > 0) {
            setFile(files[0]);
        }
    });

    function showStatus(type, message) {
        statusBox.className = "status-box " + type;
        statusBox.innerHTML = message;
    }

    convertBtn.addEventListener("click", async () => {
        const file = selectedFile;
        if (!file) return;

        const repeatFormatEl = document.querySelector('input[name="repeatFormat"]:checked');
        const repeatFormat   = repeatFormatEl ? repeatFormatEl.value : "4";
        const includeCode    = document.getElementById("includeCode").checked;
        const boldHints      = document.getElementById("boldHints").checked;

        convertBtn.disabled = true;
        showStatus('loading', '<span class="spinner"></span> ' + UI[currentLang].converting);
        await new Promise(r => setTimeout(r, 80));

        const formData = new FormData();
        formData.append("file", file, file.name);
        formData.append("repeat_format", repeatFormat);
        formData.append("include_code", includeCode ? "1" : "0");
        formData.append("bold_hints",   boldHints   ? "1" : "0");
        formData.append("language",     currentLang);

        try {
            const response = await fetch("/convert", { method: "POST", body: formData });
            const result   = await response.json();

            if (result.success) {
                const dlUrl = '/download/' + result.filename;
                let html = UI[currentLang].success + '<br>'
                         + '<a class="download-btn" href="' + dlUrl + '">' + UI[currentLang].download + '</a>';

                if (result.errors && result.errors.length > 0) {
                    // Separate guide reference, PDF notice, and real errors
                    const guideRef   = result.errors.find(e => e.startsWith('[GUIDE_REF]'));
                    const pdfNotice  = result.errors.find(e => e.startsWith('\u2139\uFE0F'));
                    const realErrors = result.errors.filter(e => !e.startsWith('[GUIDE_REF]') && !e.startsWith('\u2139\uFE0F'));

                    html += '<div style="margin-top:1rem;padding:0.75rem 1rem;background:#fef9c3;border:1px solid #fde68a;border-radius:8px;text-align:left;">';

                    // PDF notice (blue info style)
                    if (pdfNotice) {
                        html += '<div style="margin-bottom:0.75rem;padding:0.5rem 0.75rem;background:#eff6ff;border-left:3px solid #3b82f6;border-radius:4px;font-size:0.8rem;color:#1e40af;">' + pdfNotice + '</div>';
                    }

                    // Issue count and numbered list
                    if (realErrors.length > 0) {
                        html += '<strong style="color:#854d0e;">⚠ ' + realErrors.length + ' ' + UI[currentLang].warnings_post + '</strong>'
                              + '<ol style="margin:0.5rem 0 0.5rem 1.2rem;color:#854d0e;font-size:0.82rem;line-height:1.7;">';
                        realErrors.forEach(function(err) { html += '<li>' + err + '</li>'; });
                        html += '</ol><p style="margin:0.5rem 0 0;font-size:0.82rem;color:#854d0e;font-style:italic;">Kindly effect these corrections on the Word document questionnaire and re-upload.</p>';
                    }

                    // Guide reference at bottom — not numbered, distinct style
                    if (guideRef) {
                        html += '<div style="margin-top:0.75rem;padding:0.5rem 0.75rem;background:#f0f4ff;border-left:3px solid #4472C4;border-radius:4px;font-size:0.78rem;color:#4472C4;font-style:italic;">'
                              + guideRef.replace('[GUIDE_REF] ', '') + '</div>';
                    }

                    html += '</div>';
                }
                showStatus('success', html);
            } else {
                showStatus("error", UI[currentLang].error_pre + " " + result.error);
            }
        } catch (err) {
            showStatus("error", UI[currentLang].error_pre + " Something went wrong. Please try again.");
        }
        convertBtn.disabled = false;
    });

    }); // end DOMContentLoaded

    // ── Review Modal ──────────────────────────────────────────────
    let selectedRating = 0;

    function openReview() {
        selectedRating = 0;
        document.querySelectorAll('.star').forEach(s => s.textContent = '☆');
        document.getElementById('reviewComment').value = '';
        document.getElementById('reviewStatus').textContent = '';
        const modal = document.getElementById('reviewModal');
        modal.style.display = 'flex';
    }

    function closeReview() {
        document.getElementById('reviewModal').style.display = 'none';
    }

    document.querySelectorAll('.star').forEach(star => {
        star.addEventListener('mouseover', function() {
            const val = parseInt(this.dataset.val);
            document.querySelectorAll('.star').forEach((s, i) => {
                s.textContent = i < val ? '★' : '☆';
            });
        });
        star.addEventListener('mouseout', function() {
            document.querySelectorAll('.star').forEach((s, i) => {
                s.textContent = i < selectedRating ? '★' : '☆';
            });
        });
        star.addEventListener('click', function() {
            selectedRating = parseInt(this.dataset.val);
            document.querySelectorAll('.star').forEach((s, i) => {
                s.textContent = i < selectedRating ? '★' : '☆';
                s.style.color = i < selectedRating ? '#f59e0b' : '#9ca3af';
            });
        });
    });

    async function submitReview() {
        if (selectedRating === 0) {
            document.getElementById('reviewStatus').textContent = 'Please select a star rating.';
            document.getElementById('reviewStatus').style.color = '#dc2626';
            return;
        }
        const comment = document.getElementById('reviewComment').value.trim();
        try {
            const res = await fetch('/review', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ rating: selectedRating, comment })
            });
            const data = await res.json();
            if (data.success) {
                document.getElementById('reviewStatus').textContent = '✅ Thank you for your review!';
                document.getElementById('reviewStatus').style.color = '#166534';
                setTimeout(closeReview, 1500);
            } else {
                document.getElementById('reviewStatus').textContent = '❌ ' + (data.error || 'Something went wrong.');
                document.getElementById('reviewStatus').style.color = '#dc2626';
            }
        } catch(e) {
            document.getElementById('reviewStatus').textContent = '❌ Could not submit review. Please try again.';
            document.getElementById('reviewStatus').style.color = '#dc2626';
        }
    }

    // Close modal on background click
    document.getElementById('reviewModal').addEventListener('click', function(e) {
        if (e.target === this) closeReview();
    });
</script>
</body>
</html>"""


def render_page_with_lang(html, lang):
    """
    Injects a language bootstrap <script> into a page.
    The script reads lang from the URL, applies translations,
    and rewrites all internal links to preserve the lang parameter.
    """
    if lang not in ('en', 'fr'):
        lang = 'en'
    inject = f'''<script>
(function(){{
  var lang = {repr(lang)};
  window.__pageLang = lang;
  if (lang !== 'fr') return;
  document.addEventListener('DOMContentLoaded', function() {{
    applyPageLang(lang);
    // Rewrite all internal links to preserve lang=fr
    document.querySelectorAll('a[href]').forEach(function(a) {{
      var h = a.getAttribute('href');
      if (h && h.startsWith('/') && !h.includes('lang=')) {{
        a.setAttribute('href', h + (h.includes('?') ? '&' : '?') + 'lang=fr');
      }}
    }});
  }});
}})();
</script>'''
    # Inject just before </head>
    return html.replace('</head>', inject + '\n</head>', 1)

@app.route('/')
def index():
    return Response(HTML_PAGE, mimetype='text/html')

# .doc and .pdf conversion handled in convert.py load_document()



@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file was uploaded. Please select a file and try again.'})
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'No file was selected. Please choose a .docx, .doc, or .pdf file.'})
    if not allowed_file(file.filename):
        return jsonify({'success': False, 'error': 'Unsupported file type. Please upload a .docx, .doc, or .pdf file.'})
    try:
        unique_id     = str(uuid.uuid4())[:8]
        original_name = os.path.splitext(file.filename)[0]
        ext           = file.filename.rsplit('.', 1)[1].lower()
        input_filename = f'{original_name}_{unique_id}.{ext}'
        input_path     = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
        file.save(input_path)

        output_filename = f'{original_name}_{unique_id}_XLSForm.xlsx'
        output_path     = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        repeat_format   = int(request.form.get('repeat_format', 4))
        include_code    = request.form.get('include_code', '1') == '1'
        bold_hints      = request.form.get('bold_hints',   '1') == '1'
        language        = request.form.get('language', 'en')
        output_path, errors = convert(
            filepath=input_path, output_path=output_path,
            use_ai=False, repeat_format=repeat_format, include_code=include_code,
            bold_hints=bold_hints, language=language
        )
        # Post-conversion cleanup — failure here is non-fatal
        # XLSForm already created successfully at this point
        try:
            if os.path.exists(input_path):
                os.remove(input_path)
        except Exception:
            pass  # orphaned temp file acceptable — user gets their XLSForm
        # Use the actual filename from the path returned by convert()
        # in case convert() modified the output path
        actual_filename = os.path.basename(output_path)
        return jsonify({'success': True, 'filename': actual_filename, 'errors': errors})
    except Exception as e:
        # Clean up all temporary files on failure
        try:
            if 'input_path' in locals() and input_path and os.path.exists(input_path):
                os.remove(input_path)
        except Exception:
            pass
        return jsonify({'success': False, 'error': f'Conversion failed: {str(e)}'})

@app.route('/repeat-guide')
def repeat_guide():
    lang = request.args.get('lang', 'en')
    return Response(render_page_with_lang(REPEAT_GUIDE_PAGE, lang), mimetype='text/html')


@app.route('/patterns')
def patterns_guide():
    lang = request.args.get('lang', 'en')
    return Response(render_page_with_lang(PATTERNS_PAGE, lang), mimetype='text/html')


@app.route('/phone-countries')
def phone_countries():
    lang = request.args.get('lang', 'en')
    return Response(render_page_with_lang(PHONE_COUNTRIES_PAGE, lang), mimetype='text/html')


@app.route('/review', methods=['GET', 'POST'])
def review():
    import json
    import threading
    from datetime import datetime

    reviews_file = os.path.join(os.path.dirname(__file__), 'reviews.json')

    # ── Thread lock — prevents race condition on concurrent submissions ────────
    # One lock per process; ensures only one thread reads/writes at a time
    if not hasattr(review, '_lock'):
        review._lock = threading.Lock()

    if request.method == 'POST':
        # ── Bug 1 fix: guard against None from get_json() ─────────────────────
        data = request.get_json(silent=True)
        if not data or not isinstance(data, dict):
            return jsonify({
                'success': False,
                'error': 'Invalid request. Please send a JSON body with rating and comment.'
            }), 400

        try:
            rating = int(data.get('rating', 0))
        except (ValueError, TypeError):
            return jsonify({'success': False, 'error': 'Rating must be a number.'}), 400

        comment = str(data.get('comment', '')).strip()[:500]

        if rating < 1 or rating > 5:
            return jsonify({'success': False, 'error': 'Rating must be between 1 and 5.'}), 400

        # ── Bug 2 fix: atomic read-modify-write under lock ────────────────────
        with review._lock:
            reviews = []
            if os.path.exists(reviews_file):
                try:
                    with open(reviews_file, 'r') as f:
                        reviews = json.load(f)
                    if not isinstance(reviews, list):
                        reviews = []
                except (json.JSONDecodeError, IOError):
                    reviews = []

            reviews.append({
                'rating': rating,
                'comment': comment,
                'date': datetime.utcnow().strftime('%d %b %Y'),
            })

            with open(reviews_file, 'w') as f:
                json.dump(reviews, f, indent=2)

        return jsonify({'success': True})

    # ── GET — return all reviews (under lock to prevent read/write race) ────────
    with review._lock:
        reviews = []
        if os.path.exists(reviews_file):
            try:
                with open(reviews_file, 'r') as f:
                    reviews = json.load(f)
                if not isinstance(reviews, list):
                    reviews = []
            except (json.JSONDecodeError, IOError):
                reviews = []

    avg = round(sum(r['rating'] for r in reviews) / len(reviews), 1) if reviews else 0
    return jsonify({'reviews': reviews, 'average': avg, 'total': len(reviews)})


@app.route('/download/<filename>')
def download_file(filename):
    safe_filename = os.path.basename(filename)
    file_path     = os.path.join(app.config['OUTPUT_FOLDER'], safe_filename)
    if not os.path.exists(file_path):
        return jsonify({'success': False, 'error': 'File not found.'}), 404
    return send_file(file_path, as_attachment=True, download_name=safe_filename)

if __name__ == '__main__':
    print('Starting KoboToolbox XLSForm Converter...')
    print('Open your browser and go to: http://localhost:5000')
    app.run(debug=True)
