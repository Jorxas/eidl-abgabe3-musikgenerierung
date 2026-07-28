# Musikgenerierung mit RNNs

Abgabe 3 aus dem Kurs "Einführung in Deep Learning" (SoSe26). Wir – Jordan Pokem und Leslie Tsafack – haben ein LSTM (und zum Vergleich ein GRU) trainiert, das irische Volksmusik lernt und danach selbst neue Melodien schreibt.

Die Grundidee dahinter, die uns beim Bearbeiten selbst erst richtig klar wurde: Musik in ABC-Notation ist einfach nur Text. Damit wird "Musik komponieren" zu einem ganz normalen Sprachmodellierungs-Problem, wie wir es aus der Vorlesung zu RNNs und Sprache kannten – das Netz lernt Zeichen für Zeichen, welches Zeichen als nächstes kommt.

## Worum geht's

Datensatz ist [IrishMAN](https://huggingface.co/datasets/sander-wood/irishman), eine Sammlung irischer Volkslieder in ABC-Notation (das ist ein Textformat für Musik, ungefähr so: `X:1 L:1/8 M:6/8 K:D |: A | f>ga gfe | ...`). Wir haben daraus:

1. einen Zeichen-Tokenizer gebaut (Text zu Zahlen-IDs und zurück)
2. ein Modell aus Embedding, LSTM (bzw. GRU) und einer linearen Ausgabeschicht trainiert, das jeweils das nächste Zeichen vorhersagt
3. das Training mit wandb geloggt (Loss, Top-1- und Top-5-Accuracy)
4. das Modell danach Zeichen für Zeichen neue Stücke schreiben lassen (Sampling aus der Softmax-Verteilung, nicht einfach das wahrscheinlichste Zeichen – sonst wiederholt sich alles sehr schnell)

Dazu noch drei Bonusaufgaben: ein Vergleich LSTM gegen GRU, ein kleiner Regel-Check, ob die generierten Stücke überhaupt gültige ABC-Syntax haben, und eine Funktion, die ein generiertes Stück als Sinuston-WAV hörbar macht.

## Ergebnisse

Test-Set nach 10 Epochen, 8000 Stücke aus dem Datensatz, `hidden_size=256`, `seq_len=100`:

| Modell | Test-Loss | Top-1-Accuracy | Top-5-Accuracy |
|--------|-----------|-----------------|-----------------|
| LSTM   | 1,08      | 67,6 %          | 93,4 %          |
| GRU    | 1,06      | 68,5 %          | 93,6 %          |

Zum Einordnen: reines Raten läge bei einer Vokabulargröße von 92 Zeichen bei ungefähr 1 %. Das GRU war in unseren Läufen minimal besser und gleichzeitig etwas schneller trainiert – bei so kleinen Unterschieden würde man das aber nicht überbewerten, dafür müsste man mehrmals mit unterschiedlichen Seeds trainieren.

Die generierten Stücke haben in unserem Validity-Check (Kopfzeilen, Tonart, genug Taktstriche, Noten vorhanden, sauberes Ende) durchgehend alle Regeln erfüllt. Klingen tut es trotzdem eher nach "irisch angehauchtes Zufallsstück" als nach einer richtigen Melodie – wofür 10 Epochen auf einem Teil des Datensatzes auch ehrlich gesagt nicht reichen.

Die vollständigen Trainingskurven liegen in wandb (Projekt `eidl-thm/4.block_Jordan_Pokem_Leslie_Tsafack_RNN`), statische Versionen davon liegen als PNG in [util/](util/).

## Womit wir zu kämpfen hatten

Der größte Hebel war die Abwägung zwischen Datenmenge und Trainingszeit (steuerbar über `NUM_TUNES` und die Schrittweite des gleitenden Fensters) – mit mehr Daten und mehr Epochen wird es sicher deutlich besser, dauert auf der CPU aber auch entsprechend länger. Und Sampling statt Argmax bei der Generierung war wichtiger als gedacht: mit Argmax bleibt das Modell schnell in Wiederholungen hängen.

## Aufbau des Repos

```
ubung03.ipynb          Hauptnotebook: Daten, Modell, Training, Evaluation, Generierung, Bonusaufgaben
onepager_ubung03.pdf    Ein-Seiten-Zusammenfassung für die Abgabe
util/
  erstelle_onepager.py       baut den One-Pager aus ergebnisse.json + beispiele.txt + den PNGs
  erstelle_praesentation.py  baut die PowerPoint-Präsentation
  notizen_hinzufuegen.py     ergänzt die Referentennotizen in der Präsentation
  ergebnisse.json             alle Metriken und die Konfiguration aus dem Notebook-Lauf
  beispiele.txt                die generierten Musikstücke als Text
  kurven_loss.png, kurven_accuracy.png, vergleich_lstm_gru.png   Plots aus dem Notebook
  praesentation_ubung03.pptx, praesentation_ubung03_MIT_NOTIZEN.pptx
```

Trainierte Modellgewichte, die generierte WAV-Datei, der lokale wandb-Cache sowie die Vorlesungsfolien und das Aufgabenblatt des Kurses sind bewusst nicht im Repo (siehe [.gitignore](.gitignore)) – die Gewichte und die Audiodatei lassen sich mit dem Notebook jederzeit neu erzeugen, und die Kursunterlagen sind nicht unsere eigene Arbeit.

## Selbst ausführen

```bash
pip install torch datasets wandb python-pptx numpy matplotlib
```

Danach `ubung03.ipynb` von oben nach unten durchlaufen lassen. Für das wandb-Logging braucht man einen (kostenlosen) Account und `wandb login` einmal in der Umgebung – ohne Account bricht die Zelle mit der `wandb.init()`-Aufruf ab, der Rest des Notebooks funktioniert aber unabhängig davon.

Wenn das Notebook einmal komplett durchgelaufen ist, liegen `ergebnisse.json`, `beispiele.txt` und die PNGs im Arbeitsverzeichnis, und man kann aus `util/` heraus noch

```bash
python erstelle_onepager.py
python erstelle_praesentation.py
```

laufen lassen, um den One-Pager bzw. die Präsentation neu zu bauen.
