# Comment automatiser la production d'une présentation PPTX en Python?

## Contexte

En tant que scientifique des données, la capacité à mener des analyses complexes est indéniablement cruciale. Toutefois, tout aussi important est le talent de communiquer ces résultats. Une analyse poussée perd de sa valeur si elle ne peut être transmise ou comprise par les parties prenantes. Les présentations PowerPoint se présentent comme l'outil couramment utilisé pour faciliter cette communication et le partage de connaissances. Néanmoins, la création manuelle de ces présentations peut s'avérer fastidieuse et chronophage. Dans ce tutoriel, nous plongerons dans l'automatisation de ce processus en exploitant le potentiel du module python-pptx. Nous allons découvrir comment chaque élément d'une présentation PowerPoint peut être manipulé comme un objet à travers le code Python. Afin de suivre ce tutoriel, il est nécessaire d’avoir une connaissance de base de la programmation Python (comprendre les boucles, conditions, fonctions etc.), un environnement de développement intégré (IDE) fonctionnel et évidemment une connaissance de base de Power Point. 

## La thématique
Le module python-pptx offre une solution puissante pour créer et personnaliser des présentations Power Point à partir de Python. Il permet aux scientifiques des données de générer des documents visuels dynamiques et informatifs, en automatisant la création de diapositives, l'ajout de textes, d'images, et même de graphiques. Ceci est particulièrement utile pour diminuer les tâches fastidieuses et répitives que la création d'un Power Point peut présenter. Un exemple est de minutieusement placer des images aux bons endroits sur plusieurs diapositives en utilisant les lignes de guides de Power Point. Un autre exemple est la création de présentation de résultats hebdomadaires ou que la structure reste pareil mais que les données changent. Cette librairie ouvre de nouvelles possibilitées pour rendre la communication des résultats de la science des données plus rapide, efficace et reproductible.



### Installation de python-pptx
Tout d’abord, débutons par l’installation de python-pptx. 
L’installation peut se faire de plus plusieurs façons : par l’utilisation de la ligne de commande de votre environnement, le terminal (CMD) de votre ordinateur ou même l’interface utilisateur de votre environnement. Afin de simplifier le tutoriel et d’accommoder les lecteurs ayant moins d’expérience en informatique en général, seule l’installation par l’utilisation de l’interface utilisateur est détaillé. Pour plus de détails sur comment procéder pour les deux autres méthodes, il existe une multitude de ressources en ligne vous permettant de suivre les étapes à effectuer, par exemple: https://www.jetbrains.com/help/pycharm/installing-uninstalling-and-upgrading-packages.html. (L’installation de «pip» peut être un prérequis pour ces méthodes)

#### Installation par l’utilisation de l’interface utilisateur de l’environnement

  * Dirigiez-vous vers l’environnement de votre choix (dans le cas du cours, Pycharm est utilisé)
  
  * Sélectionner « Réglages » dans le menu « Fichiers » en haut à gauche
  
  * Sélectionner votre projet et puis « Interpréteur de Python »
  
  * Sélectionner le « + », entrez « python-pptx » dans la boite de recherche et installez


Maintenant que python-pptx est installé vous pouvez commencer à créer des Power Point à partir de Python!


### 3.2 Présentation

#### Création et ouverture d’une présentation 

L'objectif de cette section est de créer une première présentation vide ou d'en ouvrire une existante pour ensuite lui amener des modification. Il faut d'abord importer le module presentation. La fonction « Présentation() » de cu module cré une nouvelle présentation. Dans l'exemple suivant, cette présentation est stocké sous la variable « pres ». Cette nouvelle présentation est vierge et sans diapositive par défaut. Pour ouvrir une présentation existante, il suffit de mettre le nom de la présentation entre les parenthèses:
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
Cette ligne vient sauvegarer la présentation dans le dossier ou se trouve le projet. Il est aussi possible de spéficier ou vous voulez sauvegarder votre présentation en utilisant le chemin vers cet endroit. Ce tutoriel considère que cette connaissance est une connaissance de base déjà acquise. 

Maintenant qu'une présentation est crée, ouverte, et nous sommes en mesure de la sauvegarder, nous pouvons venir la modifier. 

#### Diapositives  
Considérons une nouvelle présentation vierge, sans dispositive, par exemple, la présentation « pres » créée par la ligne 2 du précédent exemple. Avant d’ajouter une diapositive, il faut d’abord considérer la mise en page de celle-ci. Ceux qui sont familiers avec Power Point savent qu’il existe plusieurs types de mise en page, par exemple: un titre centré avec sous-titre, un entête avec une zone de texte, une diapositive vierge etc.  Dans le module python-pptx, il existe 11 différente mise en pages que l’ont peut appeler en utilisant la propriété « slide_layouts ». Les mises en pages sont numérotées de 0 à 10; une référence sur retrouve dans les images de ce tutoriel pour illustrer quel mise en page est associée avec quel identifiant. 

Une fois que la mise en page est choisie, la diapositive peut être ajoutée à la présentation en utilisant la propriété « add_slide ». Le prochain exemple, démontre comment ajouter deux diapositives ayant deux mise en pages différentes. 
```
PageTitre_layout = pres.slide_layouts[0]
PageTitre_slide = pres.slides.add_slide(PageTitre_layout)

PageContenu_layout = pres.slide_layouts[1]
PageContenu_slide = pres.slides.add_slide(PageContenu_layout)
```

Les indices 0 et 1  de la propriété « slide_layout » indiquent le type de mise en page sélectionné. Les résultats sont stockés sous les variables terminant par « layout ». Il est particulièrement utile d'utiliser des variables pour faire référence au différents type de mise en pages surtout pour les présentations longues. La proriété « add_slide() » passe le type de mise en page en argument pour finalement ajouter cette diapositive à la présentattion.

Il est aussi possible de modifier la mise ne page d'une diapositive existante. 
```
    pres = Presentation('nom_de_la_presentation.pptx')

    diapositive_a_modifier = pres.slides[0]

    PageTitre_layout_modifie = pres.slide_layouts[2]

    diapositive_a_modifier.layout = PageTitre_layout_modifie

```
À la première ligne, on ouvre la présentation crée dans un précédent exemple. La deuxième ligne définie récupère la première diapositive de la présentation et la stock sous une nouvelle variable. La troisième, définie la mise en page désirée et la dernière est celle qui procède a la modification en tant que tel. Évidemment, il est important de sauvegarder la présentation è la suite des modification tel que montré dans la sectio  précédente. 

Bien que ce tutoriel vise à exlorer les fonctionnalités principales, il est important de savoir que cette librairie offre aussi la possibilité de créer un masque de dispositive. Plus de détails peuvent être trouvés dans la documentation liée à la fin du tutoriel. 

#### Éléments
À cet étape, nous sommes en mesure de venir ajouter de éléments aux diapositives d'un Power Point. Ces éléments incluent des formes (carrés, cercle etc.), des graphiques, des tables, des images, du texte etc. Ce tutoriel abordera les grandes lignes et les éléments les plus important, mais il est important de mentionner que cette librairie offre beaucoup plus de possibilités.

 ##### Formes
Les formes sont les rectangles, cercles, étoile par exemple que nous pouvons ajouter dans Power Point. Avec l'ajout d'autres modules, il est possible d'ajouter 180 différentes formes qui peuvent, pour la majorité, peuvent être modifié en longeur et largeur et aussi de modifier la couleur de la forme. Pour ajouter une forme à une diapositive il faut:

  * Importer le module MSO_SHAPE de  pptx.enum.shapes pour avoir accès aux différentes formes
  * Importer le module Cm de pptx.util pour facilité le choix des dimensions
  * Importer le module RBGColor de pptx.dml.color pour modifier la couleur de la forme
  * Définir ou on veut placer la forme
  * Définir la taille de la forme
  * Définir tout autre paramètre désiré comme la couleur ou la bordure

(À noter: faire référence à la section précédente pour l'installation d'une librairie si elles ne sont pas déjà installées
De plus, une multitudes de librairies et modules existes pour les formes et les couleurs, ceux présentés dans l'exemple suivant ne sont que des exemples, mais ils sont assez intuitifs et permettent beaucoup de flexibilité)
Cet exemple demontre comment ajouter un rectangle rouge au coins arrondis en haut à gauche de la diapositive
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

    pres.save('test.pptx')

```
Nous avons défini certaines caractéristiques de notre forme, mais il en existe bien d'autre comme la couleur de la bordure, l'épaisseur de la bordure, l'ombre etc. Celles que nous avons omises seront attribuées aux paramètres par défaut prédéfinis. Nous ne plongerons pas plus en détail dans la logique du code, étant donné que nous avons déjà abordé quelques concepts fondamentaux plus tôt dans le tutoriel.


 ##### Graphiques
Cette librairie nous permet aussi d'ajouter des graphiques à notre présentation Power Point. D'une part, nous pouvons ajouter les données directement dans le code python en les mentionnant explicitement. D'une autre part, il est aussi possible d'utiliser les données d'un autre fichier, par exemple un fichier Excel. Ce dernier étant beaucoup plus interessant, nous allons explorer comment faire. 

Pour ajouter un graphqiue utilisant des données d'un fichier excel, il suffit de suivre les étapes suivantes:
* Importer le module CategoryChartData et de pptx.chart.data le module XL_CHART_TYPE de pptx.enum.chart
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


Bien sur, il existe beauocup d'option pour personnaliser l'esthétique du graphique. Ces propriétés peuvent être explorés davantages dans la documentation liée à la fin du tutoriel. 
 


 Documentation: https://python-pptx.readthedocs.io/en/latest/api/slides.html#slidemasters-objects 
