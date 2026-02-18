# Omgevingswet Archivering STTR

Een Python-tool voor het archiveren van gegevens uit de Registratie Toepasbare Regels (RTR) voor Nederlandse overheden.

## Overzicht

Deze applicatie biedt scripts om gegevens uit de RTR API op te vragen en lokaal te archiveren. De tool kan:

- Een Excel-overzicht genereren van activiteiten met bijbehorende werkzaamheden en regelbeheerobjecten
- DMN-logica uit de STTR downloaden en archiveren
- Werkingsgebieden per activiteit exporteren

![Voorbeeld voor Hoogheemraadschap de Stichtse Rijnlanden](data/xlsx_example.PNG)
## Vereisten

- Python 3.7 of hoger
- API-keys voor de RTR API (zie Setup)

## Setup

### 1. Installeer dependencies

Installeer de benodigde Python-packages:

```bash
pip install -r requirements.txt
```

Of handmatig:

```bash
pip install requests xlsxwriter pandas openpyxl
```

### 2. Configureer API-keys

Vraag API-keys aan voor de productie- en pre-omgeving via het [IPLO Ontwikkelaarsportaal](https://aandeslagmetdeomgevingswet.nl/ontwikkelaarsportaal/api-register/api/omgevingsdocument-toepasbaar-opvragen/).

Maak vervolgens de volgende bestanden aan en plaats de keys daarin:

```
data/prod_API_key.txt
data/pre_API_key.txt
```

**Let op:** Deze bestanden worden automatisch uitgesloten van version control via `.gitignore`.

## Gebruik

### Basisgebruik

```bash
python code/main.py
```

### Command-line opties

| Optie | Beschrijving | Standaard |
|-------|--------------|-----------|
| `--overheid` | Naam van de overheid (gebruik underscores voor spaties) | `Hoogheemraadschap_De_Stichtse_Rijnlanden` |
| `--env` | Omgeving: `prod` of `pre` | `prod` |
| `--date` | Datum in formaat DD-MM-YYYY | Vandaag |
| `--sttr` | Archiveer DMN-logica per activiteit | Uit |
| `--location` | Archiveer werkingsgebieden per activiteit | Uit |

### Voorbeelden

Standaard uitvoering voor Hoogheemraadschap De Stichtse Rijnlanden:
```bash
python code/main.py
```

Voor een ander waterschap:
```bash
python code/main.py --overheid Waterschap_Vechtstromen
```

Met specifieke datum en pre-omgeving:
```bash
python code/main.py --overheid Wetterskip_Fryslân --env pre --date 03-03-2024
```

Volledig archief inclusief STTR-bestanden en werkingsgebieden:
```bash
python code/main.py --env prod --date 12-04-2024 --sttr --location
```

**Opmerking:** De volgorde van de flags maakt niet uit. Als een optie niet wordt opgegeven, wordt de standaardwaarde gebruikt.

## Projectstructuur

```
.
├── code/
│   ├── __init__.py
│   ├── main.py           # Hoofdprogramma
│   ├── commands.py       # Command-line argument parsing
│   ├── rtr.py           # RTR API interactie en archivering
│   ├── excel.py         # Excel bestandsgeneratie
│   └── powerbi.py       # PowerBI data verwerking
├── data/
│   ├── prod_API_key.txt # API-key voor productie (niet in git)
│   ├── pre_API_key.txt  # API-key voor pre-omgeving (niet in git)
│   └── *.xlsx           # PowerBI Excel bestanden
├── log/                 # Gegenereerde log bestanden (niet in git)
├── .gitignore
├── LICENSE
├── README.md
└── requirements.txt
```

## Output

De tool genereert verschillende outputs:

- **Excel-bestanden**: Status overzicht van activiteiten met werkzaamheden en wijzigingen
- **DMN-bestanden** (met `--sttr`): XML-bestanden met STTR-logica per activiteit
- **Werkingsgebieden** (met `--location`): Tekstbestand met mapping van werkingsgebieden naar activiteiten

## Licentie

Dit project is gelicenseerd onder de MIT License - zie het [LICENSE](LICENSE) bestand voor details.

## Bijdragen

Bijdragen zijn welkom! Zie [CONTRIBUTING.md](CONTRIBUTING.md) voor richtlijnen.

## Contact

Voor vragen over de API of om API-keys aan te vragen, bezoek het [IPLO Ontwikkelaarsportaal](https://aandeslagmetdeomgevingswet.nl/ontwikkelaarsportaal/api-register/api/omgevingsdocument-toepasbaar-opvragen/).
