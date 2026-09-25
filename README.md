# Irische Melodien mit LSTM und GRU generieren

Dieses Projekt entstand gemeinsam mit Leslie Tsafack im Kurs *Einführung in Deep Learning* (THM, Sommersemester 2026). Wir wollten herausfinden, ob ein Modell die Struktur irischer Volksmelodien lernen und eigene Stücke in **ABC-Notation** erzeugen kann.

ABC-Notation beschreibt Musik als Text. Wir konnten die Aufgabe deshalb als Zeichenfolge behandeln: Ein Tokenizer wandelt Zeichen in IDs um, ein LSTM oder GRU sagt das nächste Zeichen voraus, und beim Generieren wird wiederholt aus der Vorhersage gesampelt.

## Was wir umgesetzt haben

- Zeichen-Tokenizer sowie Training und Vergleich eines LSTM und eines GRU in PyTorch.
- Experiment-Tracking mit Weights & Biases und Auswertung von Loss sowie Top-1- und Top-5-Accuracy.
- Generierung neuer ABC-Stücke per Sampling, ein einfacher Syntax-Check und eine WAV-Ausgabe mit Sinustönen.

Der Datensatz ist [IrishMAN](https://huggingface.co/datasets/sander-wood/irishman). Im dokumentierten Lauf nutzten wir 8.000 Stücke, Sequenzen mit 100 Zeichen und 10 Epochen.

| Modell | Test-Loss | Top-1 | Top-5 |
| --- | ---: | ---: | ---: |
| LSTM | 1,08 | 67,6 % | 93,4 % |
| GRU | 1,06 | 68,5 % | 93,6 % |

Das GRU lag in diesem Lauf knapp vorn. Der Unterschied ist zu klein, um daraus ohne wiederholte Läufe mit verschiedenen Seeds eine allgemeine Überlegenheit abzuleiten. Die generierten Beispiele bestanden unseren einfachen ABC-Syntax-Check, klingen aber noch nicht wie sorgfältig komponierte Melodien. Die größte praktische Abwägung war Datenmenge gegen Trainingszeit; beim Generieren verhinderte Sampling außerdem viele Wiederholungen, die mit einer reinen Argmax-Auswahl auftraten.

## Im Repository

- [`ubung03.ipynb`](ubung03.ipynb) — Daten, Modelle, Training, Evaluation und Generierung.
- [`util/ergebnisse.json`](util/ergebnisse.json) und [`util/beispiele.txt`](util/beispiele.txt) — Metriken und erzeugte Beispiele.
- [`util/`](util/) — Diagramme und Skripte für One-Pager und Präsentation.
- [`onepager_ubung03.pdf`](onepager_ubung03.pdf) — einseitige Projektzusammenfassung.

## Nachvollziehen

Installiere Python sowie die im Notebook verwendeten Pakete:

```bash
pip install torch datasets wandb python-pptx numpy matplotlib
```

Öffne anschließend `ubung03.ipynb` und führe die Zellen der Reihe nach aus. Für das Experiment-Tracking ist ein Weights-&-Biases-Konto mit `wandb login` erforderlich. Modellgewichte, generierte WAV-Dateien und Kursunterlagen sind nicht Teil des Repositorys; der Lauf kann sie lokal erzeugen.
