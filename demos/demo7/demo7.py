from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
import pandas as pd

if __name__ == "__main__":
    # Creation de la presentation
    pres = Presentation()

    # Ajouter une diapositive vierge
    slide_layout = pres.slide_layouts[6]  # Choisir une disposition qui supporte les graphiques
    slide = pres.slides.add_slide(slide_layout)

   # Spécifier le chemin vers les données ( Dans ce cas, le fichier et dans le répertoire du projet dans le fichier data)
    chemin_excel = "../../data/Donnes_graphique.xlsx"

    # Lire la table Excel en utilisant la librairie pandas (une façon parmis plusieurs)
    table = pd.read_excel(chemin_excel)

    # Extraire les données
    donnees = list(zip(table['Jours'], table['Precipitations']))

    # Créer le graphique
    graphique_donnees = CategoryChartData()
    graphique_donnees.categories = [item[0] for item in donnees]
    graphique_donnees.add_series('Precipitations en mars', (item[1] for item in donnees))

    # Définir la taille du graphique et le positionner au centre de la diapositive
    largeur = pres.slide_width*0.75
    hauteur = pres.slide_height*0.75
    axe_x = (pres.slide_width/2)-largeur/2
    axe_y = (pres.slide_height/2)-hauteur/2

    graphique = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, axe_x, axe_y, largeur, hauteur, graphique_donnees
    )

    # Enregistrement de la présentation
    pres.save('demo7.pptx')

    # Réouverture de la présentation
    # pres = Presentation('demo7.pptx')