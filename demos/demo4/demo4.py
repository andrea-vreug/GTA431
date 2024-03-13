from pptx import Presentation


if __name__ == "__main__":

    # Charger la présentation
    pres = Presentation()

    # Ajout d'une diapositive à la présentation existante
    Espace_reserve_layout = pres.slide_layouts[8]
    Espace_reserve_slide = pres.slides.add_slide(Espace_reserve_layout)

    # Trouver l'index et le nom des espaces réservés
    for type in Espace_reserve_slide.placeholders:
        print('%d %s' % (type.placeholder_format.idx, type.name))

    # Enregistrement de la présentation
    pres.save('demo4.pptx')

    # Réouvrir la présentation modifiée
    #pres = Presentation('demo4.pptx')