## Excel - Change capture

This spike implements a model for capturing changes between 2 excel files.

*python ./spike04_excel_changes_record.py*
```text
New Records:
* ExcelRecordChange(sheet='Villes', index={'ID Ville': 'V021'}, change='added', old_value=None, new_value={'ID Ville': 'V021', 'Nom Ville': 'Ville 21', 'Région': 'Nouvelle-Aquitaine', 'Pays': 'France', 'Population': '123986'})
Removed Records:
* ExcelRecordChange(sheet='Entreprises', index={'ID Entreprise': 'E015'}, change='removed', old_value={'ID Entreprise': 'E015', 'Nom Entreprise': 'Entreprise 15', "Secteur d'Activité": 'Bâtiment', "Nombre d'Employés": '200', 'Ville': 'Ville 8'}, new_value=None)
Updated Records:
* ExcelRecordChange(sheet='Clients', index={'ID Client': 'C008'}, change='updated', old_value={'ID Client': 'C008', 'Nom': 'Dupont', 'Prénom': 'Claire', 'Âge': '51', 'Sexe': 'F', 'Entreprise': 'Entreprise 30', 'Ville': 'Ville 5'}, new_value={'ID Client': 'C008', 'Nom': 'Dupont', 'Prénom': 'Claire', 'Âge': '52', 'Sexe': 'F', 'Entreprise': 'Entreprise 30', 'Ville': 'Ville 4'})

```

## The latest version

You can find the latest version to ...

```bash
git clone https://github.com/FabienArcellier/spike-openpyxl-change-capture.git
```

## Usage

You can run the application with the following command

```bash
poetry install
poetry shell

# Génère un premier fichier excel urban_planning-01.xlsx, then copy as urban_planning-02.xlsx and make some changes
python spike00-generate_excel_data.py

# Load the excel file as Cell and print the content
python spike01_excel_load_cell.py
python spike02_excel_changes_cell.py # Show changes between 2 excel files

# Advanced version, load the excel file as Record and print the content (manage offset and header)
python spike03_excel_load_record.py
python spike04_excel_changes_record.py  # Show changes between 2 excel files
```

## Contributors

* Fabien Arcellier

## License

MIT License

Copyright (c) 2024-2024 Fabien Arcellier

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
