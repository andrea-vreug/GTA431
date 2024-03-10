# Comment automatiser la production d'une présentation PPTX en Python?

## Contexte


En tant que scientifique des données, la capacité à réaliser des analyses complexes est indéniablement cruciale. Toutefois, tout aussi important est le talent de communiquer ces résultats. Une analyse approfondie perd de sa valeur si elle ne peut être transmise ou comprise par les parties prenantes. Les présentations PowerPoint se révèlent être l'outil couramment utilisé pour faciliter cette communication et le partage de connaissances. Néanmoins, la création manuelle de ces présentations peut s'avérer fastidieuse et chronophage. Dans ce tutoriel, nous plongerons dans l'automatisation de ce processus en exploitant le potentiel du module python-pptx. Nous allons découvrir comment les éléments d'une présentation PowerPoint peut être manipulé comme un objet à travers le code Python. Nous allons survoler certains éléments tout en approfondissant d'autres aspects spécifiques. Afin de suivre ce tutoriel, il est nécessaire d'avoir une connaissance de base en programmation Python (comprendre les boucles, les conditions, les fonctions, etc.), un environnement de développement intégré (IDE) fonctionnel, et évidemment, une connaissance de base de PowerPoint. 

## La thématique

Le module python-pptx propose une solution puissante pour concevoir et personnaliser des présentations PowerPoint à partir de Python. Il permet aux scientifiques des données de générer des documents visuels dynamiques et informatifs en automatisant la création de diapositives, l'ajout de texte, d'images, voire même de graphiques. Cela s'avère particulièrement utile pour simplifier les tâches fastidieuses et répétitives souvent associées à la création d'une présentation PowerPoint. Un exemple concret est la précision requise pour positionner méticuleusement des images aux emplacements appropriés sur plusieurs diapositives en utilisant les lignes de guides de PowerPoint. Un autre cas d'utilisation est la réalisation de présentations hebdomadaires où la structure reste inchangée, mais les données évoluent. Cette bibliothèque offre de nouvelles opportunités pour rendre la communication des résultats de la science des données plus rapide, efficace et reproductible..



### Installation de python-pptx

Tout d'abord, commençons par l'installation de python-pptx. L'installation peut se faire de plusieurs manières : en utilisant la ligne de commande de votre environnement, le terminal (CMD) de votre ordinateur, ou même l'interface utilisateur de votre environnement. Afin de simplifier le tutoriel et d'accommoder les lecteurs ayant moins d'expérience en informatique en général, nous détaillerons uniquement l'installation par l'utilisation de l'interface utilisateur. Pour des instructions détaillées sur les deux autres méthodes, de nombreuses ressources en ligne sont disponibles, vous permettant de suivre les étapes nécessaires. Par exemple:

https://www.jetbrains.com/help/pycharm/installing-uninstalling-and-upgrading-packages.html. (L’installation de «pip» peut être un prérequis pour ces méthodes)

#### Installation par l’utilisation de l’interface utilisateur de l’environnement

  * Dirigiez-vous vers l’environnement de votre choix (dans le cas du cours, Pycharm est utilisé)
  
  * Sélectionner « Réglages » dans le menu « Fichiers » en haut à gauche
  
  * Sélectionner votre projet et puis « Interpréteur de Python »
  
  * Sélectionner le « + », entrez « python-pptx » dans la boite de recherche et installez


Maintenant que python-pptx est installé vous pouvez commencer à créer des Power Point à partir de Python!


### 3.2 Présentation

#### Création et ouverture d’une présentation 

L'objectif de cette section est de créer une première présentation vide ou d'ouvrir une existante pour y apporter des modifications. Tout d'abord, il est nécessaire d'importer le module "presentation". La fonction "Présentation()" de ce module crée une nouvelle présentation. Dans l'exemple suivant, cette présentation est stockée sous la variable "pres". Cette nouvelle présentation est vierge et ne comporte aucune diapositive par défaut. Pour ouvrir une présentation existante, il suffit de spécifier le nom de la présentation entre les parenthèses :

```
## importation du module Presentation
from pptx import Presentation

# Création d'une nouvelle présentation
pres = Presentation()

# Ouverture d'une présentation existante
pres = Presentation('nom_de_la_presentation.pptx')
```

Afin de sauvegarder toute manipulation, il est important de sauvegarder la présenation
```
pres.save('nom_de_la_presentation.pptx')
```

Cette ligne permet de sauvegarder la présentation dans le dossier où se trouve le projet. Il est également possible de spécifier l'emplacement où vous souhaitez sauvegarder votre présentation en utilisant le chemin vers cet endroit. Ce tutoriel suppose que cette connaissance est déjà acquise en tant que base.

Maintenant qu'une présentation est créée, ouverte, et que nous sommes en mesure de la sauvegarder, nous pouvons procéder à sa modification.

#### Diapositives  
Considérons une nouvelle présentation vierge, sans diapositive, par exemple, la présentation « pres » créée par la ligne 2 de l'exemple précédent. Avant d'ajouter une diapositive, il faut d'abord prendre en compte la mise en page de celle-ci. Ceux qui sont familiers avec PowerPoint savent qu'il existe plusieurs types de mises en page, par exemple : un titre centré avec sous-titre, un entête avec une zone de texte, une diapositive vierge, etc. Dans le module python-pptx, il existe 11 mises en page différentes que l'on peut appeler en utilisant la propriété « slide_layouts ». Les mises en page sont numérotées de 0 à 10 ; une référence illustrée des mises en page associées à leurs identifiants se trouve dans les images de ce tutoriel.

Une fois que la mise en page est choisie, la diapositive peut être ajoutée à la présentation en utilisant la propriété « add_slide ». L'exemple suivant démontre comment ajouter deux diapositives avec deux mises en page différentes.
```
PageTitre_layout = pres.slide_layouts[0]
PageTitre_slide = pres.slides.add_slide(PageTitre_layout)

PageContenu_layout = pres.slide_layouts[1]
PageContenu_slide = pres.slides.add_slide(PageContenu_layout)
```


Les indices 0 et 1 de la propriété « slide_layouts » indiquent le type de mise en page sélectionné. Les résultats sont stockés sous les variables se terminant par « layout ». Il est particulièrement utile d'utiliser des variables pour faire référence aux différents types de mises en page, surtout pour les présentations longues. La propriété « add_slide » prend le type de mise en page en argument pour finalement ajouter cette diapositive à la présentation.

Il est également possible de modifier la mise en page d'une diapositive existante.
```
    pres = Presentation('nom_de_la_presentation.pptx')

    diapositive_a_modifier = pres.slides[0]

    PageTitre_layout_modifie = pres.slide_layouts[2]

    diapositive_a_modifier.layout = PageTitre_layout_modifie

```
À la première ligne, nous ouvrons la présentation créée dans un précédent exemple. La deuxième ligne définit et récupère la première diapositive de la présentation, la stockant sous une nouvelle variable. La troisième ligne définit la mise en page désirée, et enfin, la dernière ligne procède à la modification proprement dite. Évidemment, il est important de sauvegarder la présentation après ces modifications, comme illustré dans la section précédente.

Bien que ce tutoriel se concentre sur l'exploration des fonctionnalités principales, il est essentiel de noter que cette bibliothèque offre également la possibilité de créer un masque de diapositive. Pour plus de détails, vous pouvez consulter la documentation liée à la fin du tutoriel.

#### Éléments

À cette étape, nous sommes en mesure d'ajouter des éléments aux diapositives d'une présentation PowerPoint. Ces éléments peuvent inclure des formes (carrés, cercles, etc.), des graphiques, des tables, des images, du texte, etc. Ce tutoriel couvrira les grandes lignes et les éléments les plus importants, mais il est crucial de noter que cette bibliothèque offre bien plus de possibilités.

 ##### Formes
Les formes sont des éléments tels que des rectangles, des cercles, des étoiles, par exemple, que nous pouvons ajouter dans PowerPoint. Avec l'ajout d'autres modules, il est possible d'intégrer jusqu'à 180 formes différentes, la plupart pouvant être modifiées en longueur et en largeur, ainsi que dans leur couleur. Pour ajouter une forme à une diapositive, il faut :

  *  Importer le module MSO_SHAPE de  pptx.enum.shapes pour avoir accès aux différentes formes
  * Importer le module Cm de pptx.util pour facilité le choix des dimensions
  * Importer le module RBGColor de pptx.dml.color pour modifier la couleur de la forme
  * Définir l'emplacement de la forme
  * Définir la taille de la forme
  * Définir tout autre paramètre désiré comme la couleur ou la bordure

(À noter : veuillez vous référer à la section précédente pour l'installation de la bibliothèque si elle n'est pas déjà installée.
De plus, une multitude de bibliothèques et de modules existent pour les formes et les couleurs. Ceux présentés dans l'exemple suivant ne sont que des exemples, mais ils sont assez intuitifs et offrent beaucoup de flexibilité.)
Cet exemple démontre comment ajouter un rectangle rouge avec des coins arrondis en haut à gauche de la diapositive.

```
    # Importation des modules
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Cm,
    from pptx.dml.color import RGBColor

    # Création de la présentation
    pres = Presentation()

    # Création de la diapositive
    diapositive_layout = pres.slide_layouts[0]
    diapositive_slide = pres.slides.add_slide(diapositive_layout)

    # Assignation de la propriété shapes à la diapositive
    formes = diapositive_slide.shapes

     # Définition de la taille
    hauteur = Cm(2)
    largeur = hauteur *2

     # Définition de l'emplacement
    axe_x = axe_y = Cm(3)

    # Ajout de la forme
    rectangle = formes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, axe_x, axe_y, largeur, hauteur
    )

    # Définition de la couleur
    remplissage = rectangle.fill
    remplissage.solid()
    remplissage.fore_color.rgb = RGBColor(255, 0, 0)

    pres.save('rectangle.pptx')

```

Nous avons défini certaines caractéristiques de notre forme, mais il en existe bien d'autres, telles que la couleur de la bordure, l'épaisseur de la bordure, l'ombre, etc. Celles que nous avons omises seront attribuées aux paramètres par défaut prédéfinis. Nous n'approfondirons pas davantage la logique du code, étant donné que nous avons déjà abordé quelques concepts fondamentaux plus tôt dans le tutoriel.


 ##### Graphiques

Ce module nous permet également d'ajouter des graphiques à notre présentation PowerPoint. D'une part, nous pouvons intégrer les données directement dans le code Python en les mentionnant explicitement. D'autre part, il est aussi possible d'utiliser les données d'un autre fichier, par exemple un fichier Excel. Cette dernière option plus intéressante, explorons comment procéder. 

Pour ajouter un graphqiue utilisant des données d'un fichier excel, il suffit de suivre les étapes suivantes:
* Importer le module CategoryChartData de pptx.chart.data et XL_CHART_TYPE de pptx.enum.chart
* Ouvrir une présentation existante ou créer une nouvelle présentation ayant une mise en page pouvant supporter un graphique
* Définir le chemin vers le fichier Excel contenant les données
* Lire les données
* Créer le graphique

Voici un exemple
```
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    import pandas as pd

    # Creation de la presentation
    pres = Presentation()

    # Ajouter une diapositive vierge
    slide_layout = pres.slide_layouts[6]  # Choose a layout that supports charts
    slide = pres.slides.add_slide(slide_layout)

    # Spécifier le chemin vers les données ( Dans ce cas, le fichier et dans le répertoire du projet)
    chemin_excel = 'Donnes_graphique.xlsx'

    # Lire la table Excel en utilisant la livrairie pandas (une façon parmis plusieurs)
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

    pres.save('precipitations_mars.pptx')
```
Bien sûr, il existe de nombreuses options pour personnaliser l'esthétique du graphique. Ces propriétés peuvent être explorées davantage dans la documentation, disponible à la fin du tutoriel
 


 Documentation: https://python-pptx.readthedocs.io/en/latest/api/slides.html#slidemasters-objects 
