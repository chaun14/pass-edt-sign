# Pass Emploi du Temps Sign

Ce projet permet de générer et de signer automatiquement des emplois du temps au format PDF pour les alternant 3A d'IMT Atlantique. L'application propose une interface graphique simple pour sélectionner la semaine souhaitée et génère un PDF signé.

## Installation

1. Clone / Download the repository:

2. Install dependencies (preferably in a virtual environment):

   ```bash
   pip install -r requirements.txt
   ```

   If you are using a debian-based system, you might need to install python packages through apt:

   ```bash
   sudo apt update
   sudo apt install -y \
    python3-requests \
    python3-selenium \
    python3-dotenv \
    python3-pil \
    python3-pypdf2 \
    python3-reportlab

   ```

## Running the GUI

To start the GUI application, run:

```bash
python gui.py
```
