## Création et ouverture d’une présentation
## Nom du fichier : demo1

## importation du module Presentation
from pptx import Presentation


if __name__ == "__main__":
    # Création d'une nouvelle présentation
    pres = Presentation()

    # Ajout de deux diapositives
    PageTitre_layout = pres.slide_layouts[0]
    PageTitre_slide = pres.slides.add_slide(PageTitre_layout)

    PageContenu_layout = pres.slide_layouts[1]
    PageContenu_slide = pres.slides.add_slide(PageContenu_layout)

    # Enregistrement de la présentation
    pres.save('demo2.pptx')

    # Ouverture d'une présentation existante
   #pres = Presentation('demo2.pptx')
