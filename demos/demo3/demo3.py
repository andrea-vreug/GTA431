from pptx import Presentation


if __name__ == "__main__":

    # Charger la présentation
    pres = Presentation('../demo2/demo2.pptx')

    # Récupérer la diapositive à modifier
    diapositive_a_modifier = pres.slides[0]

    # Obtenir la disposition (layout) originale de la diapositive
    layout_original = diapositive_a_modifier.slide_layout

    # Obtenir la disposition modifiée à assigner
    PageTitre_layout_modifie = pres.slide_layouts[2]

    # Assigner la disposition modifiée à la diapositive
    diapositive_a_modifier.layout = PageTitre_layout_modifie

    # Afficher les propriétés de la disposition originale et modifiée pour comparaison
    print("Disposition Originale :")
    print(layout_original.name)
    print("Disposition Modifiée :")
    print(PageTitre_layout_modifie.name)

    # Enregistrer la présentation modifiée
    pres.save('demo3.pptx')

    # Réouvrir la présentation modifiée
    #pres = Presentation('demo3.pptx')
