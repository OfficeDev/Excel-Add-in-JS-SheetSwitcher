# Exemple de complément de volet Office - Sélecteur de feuilles pour Excel 2016

_S’applique à : Excel 2016_

Ce complément de volet Office offre un moyen d’ajouter de nouvelles feuilles de calcul à un classeur via le volet Office et de naviguer facilement entre les feuilles de calcul dans Excel 2016. Il a deux versions : éditeur de texte et Visual Studio.

![Exemple de sélecteur de feuilles](../images/SheetSwitcher_taskpane.PNG)

## Essayez !
### Version d’éditeur de texte

Pour déployer et tester votre complément, le plus simple consiste à copier les fichiers sur un partage réseau.

1.  Créez un dossier dans un partage réseau (par exemple, \\\MyShare\SheetSwitcher) et copiez tous les fichiers dans le dossier d’éditeur de texte. 
2.  Modifiez l’élément <SourceLocation> du fichier manifeste afin qu’il pointe vers l’emplacement de partage de l’étape 1. 
3.  Copiez le fichier manifeste (SheetSwitcherManifest.xml) dans un partage réseau (par exemple, \\\MyShare\MyManifests).
4.  Ajoutez l’emplacement de partage qui contient le fichier manifeste sous forme de catalogue d’applications approuvées dans Excel.

    a. Lancez Excel et ouvrez une feuille de calcul vide.  
    
    b. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
    
    c. Choisissez **Centre de gestion de la confidentialité**, puis cliquez sur le bouton **Paramètres du Centre de gestion de la confidentialité**.
    
    d. Choisissez **Catalogues d’applications approuvées**.
    
    e. Dans la zone **URL du catalogue**, entrez le chemin d’accès du partage réseau que vous avez créé à l’étape 1, puis choisissez **Ajouter un catalogue**.
    
   Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**. Un message s’affiche pour vous informer que vos paramètres seront appliqués la prochaine fois que vous démarrerez Office. 
        
5.  Testez et exécutez le complément. 

    a. Dans l’onglet **Insertion** d’Excel 2016, choisissez **Mes compléments**. 
    
    b. Dans la boîte de dialogue **Compléments Office**, choisissez **Dossier partagé**.
    
    c. Choisissez **Exemple de complément de sélecteur de feuille**>**Insertion**. Le complément s’ouvre dans un volet Office à droite de la feuille de calcul active, comme indiqué dans l’illustration suivante. 
        
  ![Exemple de sélecteur de feuilles](../images/SheetSwitcher_taskpane.PNG)

    d. Cliquez sur le bouton **Ajouter des feuilles**. Cette opération ajoute 14 nouvelles feuilles à la feuille de calcul. Cliquez sur l’un des boutons correspondant à une feuille affichés dans le volet Office pour accéder à la feuille dans le classeur.
        

### Version de Visual Studio
1.  Copiez le projet dans un dossier local et ouvrez le fichier Excel-Add-in-Sheet-Switcher.sln dans Visual Studio.
2.  Appuyez sur F5 pour créer et déployer l’exemple de complément. Excel démarre et le complément s’ouvre dans un volet Office à droite de la feuille de calcul active, comme indiqué dans l’illustration suivante. 
        
  ![Exemple de sélecteur de feuilles](../images/SheetSwitcher_taskpane.PNG)

3. Cliquez sur le bouton **Ajouter des feuilles**. Cette opération ajoute 14 nouvelles feuilles à la feuille de calcul. Cliquez sur l’un des boutons correspondant à une feuille affichés dans le volet Office pour accéder à la feuille dans le classeur.



### En savoir plus

Les API JavaScript pour Excel peuvent vous offrir beaucoup pour l’élaboration de vos compléments. Voici quelques-unes des ressources disponibles : 

1.  [Présentation de la programmation pour les compléments Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Explorateur d’extraits de code pour Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Exemples de code pour les compléments Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [Référence de l’API JavaScript pour les compléments Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Créer son premier complément Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)
