## Création et ouverture d’une présentation
## Nom du fichier : demo1

## importation du module Presentation
from pptx import Presentation


if __name__ == "__main__":
    # Création d'une nouvelle présentation
    pres = Presentation()

    pres.save('demo1.pptx')

    # Ouverture d'une présentation existante
    #pres = Presentation('demo1.pptx')