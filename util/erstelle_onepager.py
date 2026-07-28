# -*- coding: utf-8 -*-
"""
Erstellt den One-Pager (PDF, eine A4-Seite, 12 pt) zur Abgabe 3.

Voraussetzung: Das Notebook `ubung03.ipynb` wurde einmal komplett ausgeführt
(ergebnisse.json, beispiele.txt, kurven_loss.png, kurven_accuracy.png).

Aufruf:  python erstelle_onepager.py
Ausgabe: onepager_ubung03.pdf  (+ onepager_ubung03.html als Zwischenschritt)
"""
import base64
import json
import os
import subprocess
import sys

# ----------------------------------------------------------------- Daten laden
if not os.path.exists("ergebnisse.json"):
    sys.exit("FEHLER: ergebnisse.json nicht gefunden – bitte zuerst das Notebook ausführen.")

with open("ergebnisse.json", encoding="utf-8") as f:
    erg = json.load(f)

cfg = erg["config"]
lstm, gru = erg["lstm"], erg["gru"]
zufall = 1.0 / max(erg["vocab_groesse"], 1)

with open("beispiele.txt", encoding="utf-8") as f:
    beispiel = f.read().split("\n\n")[0]

# ABC-Beispiel für den Platz auf der Seite leicht kürzen
beispiel_zeilen = [z[:86] for z in beispiel.split("\n")][:8]
beispiel_html = "\n".join(beispiel_zeilen).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def bild64(pfad):
    """PNG als data-URI einbetten, damit das HTML komplett eigenständig ist."""
    with open(pfad, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()


def pz(x):
    """Prozentwert mit deutschem Dezimalkomma (geschütztes Leerzeichen vor %)."""
    return f"{x * 100:.1f}".replace(".", ",") + "&nbsp;%"


def tausender(x):
    """Ganzzahl mit deutschem Tausenderpunkt."""
    return f"{x:,}".replace(",", ".")


# ----------------------------------------------------------------- HTML bauen
html = f"""<!doctype html>
<html lang="de"><head><meta charset="utf-8"><title>One-Pager Abgabe 3</title>
<style>
  @page {{ size: A4; margin: 0; }}
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{ width: 210mm; height: 297mm; padding: 10mm 12mm; overflow: hidden;
         font-family: "Segoe UI", sans-serif; font-size: 12pt; color: #0b0b0b; line-height: 1.3; }}
  h1 {{ font-size: 16pt; }}
  .meta {{ color: #52514e; font-size: 12pt; margin-top: 1mm; }}
  .rule {{ height: 1.2mm; width: 28mm; background: #2a78d6; margin: 2.5mm 0 4mm; }}
  h2 {{ font-size: 12pt; color: #16324f; margin-bottom: 1.5mm; }}
  section {{ margin-bottom: 3.5mm; }}
  .spalten {{ display: flex; gap: 7mm; }}
  .spalten > div {{ flex: 1; }}
  ul {{ padding-left: 5mm; }}
  li {{ margin-bottom: 0.8mm; }}
  table {{ border-collapse: collapse; width: 100%; }}
  th, td {{ border-bottom: 0.3mm solid #e1e0d9; padding: 1mm 2mm; text-align: left; font-size: 12pt; }}
  th {{ color: #52514e; font-weight: 600; }}
  .zahl {{ text-align: right; white-space: nowrap; }}
  .bilder {{ display: flex; gap: 5mm; }}
  .bilder img {{ width: 91mm; }}
  pre {{ font-family: Consolas, monospace; font-size: 10pt; background: #f0efec;
        border-radius: 2mm; padding: 2.5mm 3mm; white-space: pre-wrap; }}
  .hinweis {{ color: #52514e; }}
</style></head><body>

<h1>Musikgenerierung mit RNNs — Abgabe 3</h1>
<div class="meta">Einführung in Deep Learning (SoSe26) · Jordan Pokem &amp; Leslie Tsafack · 16.07.2026</div>
<div class="rule"></div>

<div class="spalten"><div>

<section>
<h2>Ansatz &amp; Datenvorverarbeitung</h2>
<ul>
<li>IrishMAN-Datensatz (Hugging Face); {tausender(cfg["num_tunes"])} Stücke aus dem Train-Split.</li>
<li>Zeichen-Tokenizer ({erg["vocab_groesse"]} Zeichen Vokabular): Encoding Text&nbsp;→&nbsp;IDs, Decoding IDs&nbsp;→&nbsp;Text.</li>
<li>Gleitendes Fenster (Länge {cfg["seq_len"]}, Schritt {cfg["schritt"]}) → {tausender(erg["anzahl_sequenzen"])} Sequenzen; Ziel = Input um 1 Zeichen verschoben (Next-Token-Prediction). Split 80/10/10.</li>
</ul>
</section>

<section>
<h2>Architektur &amp; Hyperparameter</h2>
<ul>
<li>Embedding ({cfg["embed_dim"]}) → LSTM ({cfg["num_layers"]}&nbsp;×&nbsp;{cfg["hidden_size"]}) → Linear (→ {erg["vocab_groesse"]} Logits); {tausender(erg["anzahl_parameter"])} Parameter.</li>
<li>CrossEntropy-Loss über alle Zeitschritte, Adam (LR {cfg["learning_rate"]}), Batch {cfg["batch_size"]}, {cfg["num_epochs"]} Epochen; Training in wandb geloggt (eidl-thm).</li>
<li>Generierung: autoregressives Sampling (Softmax + <span style="font-family:Consolas">torch.multinomial</span>) ab einem Seed-Header.</li>
</ul>
</section>

</div><div>

<section>
<h2>Ergebnisse (Test-Set)</h2>
<table>
<tr><th>Modell</th><th class="zahl">Top-1</th><th class="zahl">Top-5</th><th class="zahl">Loss</th></tr>
<tr><td>LSTM (Run&nbsp;1)</td><td class="zahl">{pz(lstm["test"]["top1"])}</td><td class="zahl">{pz(lstm["test"]["top5"])}</td><td class="zahl">{lstm["test"]["loss"]:.3f}</td></tr>
<tr><td>GRU (Bonus-Ablation)</td><td class="zahl">{pz(gru["test"]["top1"])}</td><td class="zahl">{pz(gru["test"]["top5"])}</td><td class="zahl">{gru["test"]["loss"]:.3f}</td></tr>
</table>
<p class="hinweis" style="margin-top:1.5mm">Zufallsbaseline Top-1 ≈ {pz(zufall)}. Bonus Validity-Check (6 ABC-Grammatikregeln): {erg["validity_score"] * 100:.0f}&nbsp;% erfüllt. Bonus Kreativität: Ausgabe als hörbare WAV-Datei (eigener Mini-ABC-Parser + Sinus-Synthese).</p>
</section>

<section>
<h2>Herausforderungen</h2>
<ul>
<li>Datenmenge vs. Trainingszeit → Teil-Set und Fenster-Schrittweite als Stellschrauben.</li>
<li>Argmax-Generierung wiederholt sich → Sampling aus der Softmax-Verteilung.</li>
<li>Lange Stücke verlieren Struktur (Kontext = Fensterlänge) → Abschneiden an der ersten Leerzeile.</li>
</ul>
</section>

</div></div>

<section>
<h2>Trainingskurven (LSTM)</h2>
<div class="bilder">
<img src="{bild64('kurven_loss.png')}" alt="Loss pro Epoche">
<img src="{bild64('kurven_accuracy.png')}" alt="Accuracy pro Epoche">
</div>
</section>

<section>
<h2>Beispiel: generierte Musik (ABC-Notation)</h2>
<pre>{beispiel_html}</pre>
</section>

</body></html>
"""

with open("onepager_ubung03.html", "w", encoding="utf-8") as f:
    f.write(html)

# ----------------------------------------------------------------- PDF drucken (Edge headless)
edge_kandidaten = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]
edge = next((p for p in edge_kandidaten if os.path.exists(p)), None)
if edge is None:
    sys.exit("HTML erstellt. Edge nicht gefunden – bitte onepager_ubung03.html im Browser "
             "öffnen und als PDF drucken (A4, ohne Kopf-/Fußzeile).")

html_url = "file:///" + os.path.abspath("onepager_ubung03.html").replace("\\", "/")
pdf_pfad = os.path.abspath("onepager_ubung03.pdf")
subprocess.run([edge, "--headless", "--disable-gpu", "--no-pdf-header-footer",
                f"--print-to-pdf={pdf_pfad}", html_url], check=True, timeout=120)
print(f"One-Pager gespeichert: {pdf_pfad}")
