from pptx import Presentation


if __name__ == "__main__":

    # Charger la présentation
    pres = Presentation('../demo4/demo4.pptx')

    # Spécifier la diapositive à modifier
    Espace_reserve_slide = pres.slides[0]

    # Ajouter un titre
    titre_espace = Espace_reserve_slide.placeholders[0]
    titre = "Université de Sherbrooke"
    titre_espace.text = titre

    # Ajouter un sous-titre
    soustitre_espace = Espace_reserve_slide.placeholders[2]
    soustitre = "GTA431"
    soustitre_espace.text = soustitre

    # Ajouter une image
    image_espace = Espace_reserve_slide.placeholders[1]
    chemin_image = "../../data/Logo-UDS.png"
    image_espace.insert_picture(chemin_image)

    # Enregistrement de la présentation
    pres.save('demo5.pptx')

    # Réouvrir la présentation modifiée
    #pres = Presentation('demo5.pptx')