# Importation des modules
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Cm
from pptx.dml.color import RGBColor

if __name__ == "__main__":

    # Création de la présentation
    pres = Presentation()

    # Création de la diapositive avec mise en page
    diapositive_layout = pres.slide_layouts[0]
    diapositive_slide = pres.slides.add_slide(diapositive_layout)

    # Assignation de la propriété shapes à la diapositive
    formes = diapositive_slide.shapes

    # Définition de la taille en centimètres
    hauteur = Cm(2)
    largeur = hauteur * 2

    # Définition de l'emplacement (axe x et y de la diapositive)
    axe_x = axe_y = Cm(3)

    # Ajout de la forme (important de respecter l'ordre des arguments)
    rectangle = formes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, axe_x, axe_y, largeur, hauteur
    )

    # Définition de la couleur
    remplissage = rectangle.fill
    remplissage.solid()
    remplissage.fore_color.rgb = RGBColor(255, 0, 0)

    # Enregistrement de la présentation
    pres.save('demo6.pptx')

    # Réouverture de la présentation
    #pres = Presentation('demo6.pptx')