APPLICATION SUIVI INTERVENTIONS - H2O TECH

Variables Render à ajouter :
ACCESS_CODE = h2otech
PYTHON_VERSION = 3.12.7
DATABASE_URL = Internal Database URL PostgreSQL

Build Command : pip install -r requirements.txt
Start Command : gunicorn app:app

Routes :
/ = ajout et suivi interventions
/ecran = affichage atelier
/historique?code=h2otech = historique complet
/export/monthly?code=h2otech = export Excel du mois
/export/all?code=h2otech = export Excel complet
