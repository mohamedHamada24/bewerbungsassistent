# Bewerbungsassistent

Ein KI-gestützter Bewerbungsassistent, der LinkedIn-Jobs durchsucht und automatisch ein passendes **Anschreiben** sowie einen **angepassten Lebenslauf** generiert.

## Features

- **LinkedIn-Jobsuche** – Sucht passende Stellen direkt über deinen LinkedIn-Account
- **KI-Bewertung** – Claude bewertet jede Stelle auf Übereinstimmung mit deinem Lebenslauf (Score 1–10)
- **Anschreiben-Generator** – Erstellt professionelle deutsche Anschreiben
- **Lebenslauf-Anpassung** – Passt deinen Lebenslauf an die jeweilige Stelle an
- **DOCX-Export** – Beide Dokumente als Word-Datei herunterladen
- **Manuelle Eingabe** – Stellenbeschreibung direkt einfügen möglich

## Voraussetzungen

- Python 3.10+
- [Anthropic API Key](https://console.anthropic.com) (kostenpflichtig, ~$5 reicht lange)
- LinkedIn Account (Cookie-Authentifizierung)

## Installation

```bash
git clone https://github.com/DEIN_USERNAME/bewerbungsassistent.git
cd bewerbungsassistent

python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Starten

```bash
venv/bin/streamlit run app.py
```

Die App öffnet sich unter `http://localhost:8501`.

## Einrichtung

### Anthropic API Key
1. Konto erstellen auf [console.anthropic.com](https://console.anthropic.com)
2. API Key generieren und in der App-Seitenleiste einfügen

### LinkedIn Cookie
1. LinkedIn in Chrome öffnen und einloggen
2. `F12` → **Application** → **Cookies** → `https://www.linkedin.com`
3. Werte von `li_at` und `JSESSIONID` kopieren
4. In der App-Seitenleiste einfügen und auf **💾 Speichern** klicken

## Nutzung

1. **Lebenslauf** als PDF hochladen
2. **Jobtitel** und **Ort** eingeben (z. B. `Werkstudent Informatik`, `Frankfurt`)
3. Auf **🔍 Jobs suchen & bewerten** klicken
4. Ergebnisse nach Score filtern
5. Bei passenden Stellen auf **📄 Unterlagen erstellen** klicken
6. Dokumente herunterladen und selbst bewerben

## Projektstruktur

```
bewerbungsassistent/
├── app.py              # Streamlit UI (Hauptdatei)
├── main.py             # CLI-Version (ohne UI)
├── requirements.txt    # Python-Abhängigkeiten
└── README.md
```

## Hinweis

Die LinkedIn-Authentifizierung über Cookies ist für den persönlichen Gebrauch gedacht. Die Zugangsdaten werden ausschließlich lokal gespeichert (`.config.json`) und nie übertragen.
