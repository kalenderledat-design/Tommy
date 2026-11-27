# Menen bedrijven
`extract_menen_full.py` haalt alle bedrijven in Menen (naam, adres, telefoon, e-mail, website) uit OpenStreetMap
en schrijft ze naar `data/menen_bedrijven.csv`. Een GitHub Action draait dit automatisch elke maandag.
Handmatig: `python extract_menen_full.py` of Actions → "Menen bedrijven bijwerken" → Run workflow.
