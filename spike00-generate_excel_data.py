import pandas as pd
import random

# Number of lines per sheet
n_clients = 50
n_villes = 20
n_entreprises = 30

clients_data = {
    "ID Client": [f"C{str(i+1).zfill(3)}" for i in range(n_clients)],
    "Nom": [random.choice(["Dupont", "Martin", "Lefevre", "Moreau", "Bernard"]) for _ in range(n_clients)],
    "Prénom": [random.choice(["Marie", "Jean", "Pierre", "Claire", "Luc"]) for _ in range(n_clients)],
    "Âge": [random.randint(20, 70) for _ in range(n_clients)],
    "Sexe": [random.choice(["M", "F"]) for _ in range(n_clients)],
    "Entreprise": [random.choice([f"Entreprise {i+1}" for i in range(n_entreprises)]) for _ in range(n_clients)],
    "Ville": [random.choice([f"Ville {i+1}" for i in range(n_villes)]) for _ in range(n_clients)]
}
clients_df = pd.DataFrame(clients_data)

villes_data = {
    "ID Ville": [f"V{str(i+1).zfill(3)}" for i in range(n_villes)],
    "Nom Ville": [f"Ville {i+1}" for i in range(n_villes)],
    "Région": [random.choice(["Île-de-France", "Auvergne-Rhône-Alpes", "Nouvelle-Aquitaine", "Occitanie"]) for _ in range(n_villes)],
    "Pays": ["France" for _ in range(n_villes)],
    "Population": [random.randint(5000, 1000000) for _ in range(n_villes)]
}
villes_df = pd.DataFrame(villes_data)

entreprises_data = {
    "ID Entreprise": [f"E{str(i+1).zfill(3)}" for i in range(n_entreprises)],
    "Nom Entreprise": [f"Entreprise {i+1}" for i in range(n_entreprises)],
    "Secteur d'Activité": [random.choice(["Informatique", "Bâtiment", "Santé", "Finance", "Éducation"]) for _ in range(n_entreprises)],
    "Nombre d'Employés": [random.randint(10, 500) for _ in range(n_entreprises)],
    "Ville": [random.choice([f"Ville {i+1}" for i in range(n_villes)]) for _ in range(n_entreprises)]
}
entreprises_df = pd.DataFrame(entreprises_data)

# Generate excel file
output_path = "urban_planning-01.xlsx"
with pd.ExcelWriter(output_path) as writer:
    clients_df.to_excel(writer, sheet_name="Clients", index=False)
    villes_df.to_excel(writer, sheet_name="Villes", index=False)
    entreprises_df.to_excel(writer, sheet_name="Entreprises", index=False)

