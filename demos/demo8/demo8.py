from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Cm
from pptx.dml.color import RGBColor
import pandas as pd

if __name__ == "__main__":
    # Creation de la presentation
    pres = Presentation()

    ## DIAPOSITIVE TITRE
    # Ajouter une diapositive pour le titre de la présentation
    diapo_titre_layout = pres.slide_layouts[8]
    diapo_titre_slide = pres.slides.add_slide(diapo_titre_layout)

    # Ajouter un titre
    titre_espace = diapo_titre_slide.placeholders[0]
    titre = "Université de Sherbrooke"
    titre_espace.text = titre

    # Ajouter un sous-titre
    soustitre_espace = diapo_titre_slide.placeholders[2]
    soustitre = "GTA431"
    soustitre_espace.text = soustitre

    # Ajouter une image
    image_espace = diapo_titre_slide.placeholders[1]
    chemin_image = "../../data/Logo-UDS.png"
    image_espace.insert_picture(chemin_image)



    ## DIAPOSITIVE AVEC GRAPHIQUE
    # Ajouter une diapositive vierge
    diapo_graphique_layout = pres.slide_layouts[6]
    diapo_graphique_slice = pres.slides.add_slide(diapo_graphique_layout)

    # Spécifier le chemin vers les données ( Dans ce cas, le fichier et dans le répertoire du projet dans le fichier data)
    chemin_excel = "../../data/Donnes_graphique.xlsx"

    # Lire la table Excel en utilisant la librairie pandas (une façon parmis plusieurs)
    table = pd.read_excel(chemin_excel)

    # Extraire les données
    donnees_graph = list(zip(table['Jours'], table['Precipitations']))

    # Créer le graphique
    graphique_donnees = CategoryChartData()
    graphique_donnees.categories = [item[0] for item in donnees_graph]
    graphique_donnees.add_series('Precipitations en mars', (item[1] for item in donnees_graph))

    # Définir la taille du graphique et le positionner au centre de la diapositive
    largeur_graph = pres.slide_width * 0.75
    hauteur_graph = pres.slide_height * 0.75
    axe_x_graph = (pres.slide_width / 2) - largeur_graph / 2
    axe_y_graph = (pres.slide_height / 2) - hauteur_graph / 2

    graphique = diapo_graphique_slice.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, axe_x_graph, axe_y_graph, largeur_graph, hauteur_graph, graphique_donnees
    )


    ## DIAPOSITIVE CONCLU

    diapo_conclu_layout = pres.slide_layouts[5]
    diapo_conclu_slide = pres.slides.add_slide(diapo_conclu_layout)

    titre_espace = diapo_conclu_slide.placeholders[0]
    titre = "FÉLICITATIONS!"
    titre_espace.text = titre

    # Assignation de la propriété shapes à la diapositive
    formes = diapo_conclu_slide.shapes

    # Définition de la taille en centimètres
    hauteur_etoile = Cm(8)
    largeur_etoile = hauteur_etoile

    # Définition de l'emplacement (axe x et y de la diapositive)
    axe_x_etoile = (pres.slide_width/2) -  largeur_etoile / 2
    axe_y_etoile = (pres.slide_height/2) -  hauteur_etoile / 2

    # Ajout de la forme (important de respecter l'ordre des arguments)
    etoile = formes.add_shape(
        MSO_SHAPE.STAR_5_POINT, axe_x_etoile, axe_y_etoile, largeur_etoile, hauteur_etoile  # Utilisation de MSO_SHAPE.STAR_5
    )

    # Définition de la couleur jaune
    remplissage = etoile.fill
    remplissage.solid()
    remplissage.fore_color.rgb = RGBColor(255, 255, 0)  # Jaune (255, 255, 0)



    # Enregistrement de la présentation
    pres.save('demo8.pptx')

    # Réouverture de la présentation
    #pres = Presentation('demo8.pptx')