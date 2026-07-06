# EQDS Word To XLSForm Converter

A web-based tool that automatically converts Word document questionnaires into KoboToolbox-ready XLSForm Excel files.

Built following **The Enubiaka Questionnaire Design Standard (EQDS)** — a structured questionnaire design system that ensures consistent, high-quality XLSForm output.

---

## Features

- 📄 Upload any `.docx` questionnaire and get a ready-to-use XLSForm
- 🔁 Five repeat group formats including Smart Auto-Select
- 📊 Full support for all KoboToolbox question types
- 🧮 Automatic SEC (Socio-Economic Class) score calculations
- 🔗 Reference call resolution across questions
- ⚠️ Plain English error reporting with fix suggestions
- 🌍 Bilingual support — English and French questionnaires
- 🎨 Clean, professional web interface

## Supported Question Types

- Single choice / Multiple choice
- Open ended (text, integer, decimal)
- Grids (single choice, multiple choice, open ended)
- Semantic differential grids (bipolar scale)
- MaxDiff grids
- PSM / Van Westendorp price sensitivity meter
- Funnel questions
- SEC questions with score calculations
- Note questions
- Geopoint, date, time, image, audio, video, barcode and more

## Repeat Group Formats

| Format | Best For |
|---|---|
| Automated Repeat Loop | Large/variable lists, household surveys |
| Brand-Fixed Group Loop | Brand equity, competitive benchmarking |
| Sequential Positional Loop | Product testing, ranked preference |
| Smart Auto-Select | Engine decides per repeat group |
| Direct Scripting Mode | No system repeat groups |

## Tech Stack

- Python 3.12
- Flask
- python-docx
- pandas
- openpyxl

## Access
The tool is live at [eqdsword2xlsform.com](https://eqdsword2xlsform.com) — no installation required.

For local development, contact the author.

## Deployment
Deployed on Render at [eqdsword2xlsform.com](https://eqdsword2xlsform.com).
Compatible with any platform supporting Python/Flask.

```
web: gunicorn app:app
```

---

*By Marvis Onyenwenu Enubiaka*
