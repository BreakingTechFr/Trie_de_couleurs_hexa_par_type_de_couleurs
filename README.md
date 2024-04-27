# Tri de couleurs by BreakingTech

Ce script Python utilise la bibliothèque tkinter pour créer une interface graphique permettant de trier et d'afficher par type de couleurs, des couleurs hexadecimales à partir d'un fichier Excel.

## Utilisation

1. Assurez-vous d'avoir installé les dépendances requises en exécutant `pip install -r requirements.txt`.
2. Exécutez le script en exécutant la commande `python script.py`.
3. Cliquez sur le bouton "Import Excel" pour sélectionner un fichier Excel contenant les données de couleur.
4. Les couleurs seront triées et affichées dans l'interface graphique.
5. Vous pouvez également choisir d'afficher les informations supplémentaires telles que le HUE, le RGB et le pourcentage de saturation en cliquant sur les boutons correspondants.
6. Les données triées peuvent être exportées au format CSV en cliquant sur le bouton "Exporter CSV".

## Remarque

Ce script nécessite l'installation de la bibliothèque Pandas version 1.3.3. Vous pouvez installer cette bibliothèque en exécutant `pip install pandas==1.3.3`.
Un fichier execel d'exemple est donné dans ce repo sous le nom : excel.xlsx