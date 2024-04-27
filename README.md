# Trie de couleurs hexa par type de couleurs

![Capture d’écran 2024-04-27 à 17 35 13](https://github.com/BreakingTechFr/Trie_de_couleurs_hexa_par_type_de_couleurs/assets/128238555/6339b21d-7447-416b-a3f0-33a3c0f97baf)

Ce script Python permet de trier et d'afficher par type de couleurs, des couleurs hexadecimales à partir d'un fichier Excel.
Je l'ai créé notamment pour trier des feutres de la marque "Ohuhu".
Pour l'utiliser, il faut créer un fichie excel comportant 3 colonnes : Code (le code fabricant du feutre), Nom (nom de la couleur) et Couleur(la couleur hexadecimale du feutre).
Un fichier execel d'exemple est donné dans ce repo qui m'a servi à trier les "Ohuhu Stylo Marqueur, 320 Couleurs Double Pointe", sous le nom : excel.xlsx

## Utilisation

1. Assurez-vous d'avoir installé les dépendances requises en exécutant :
```shell
pip install -r requirements.txt
```
2. Exécutez le script en exécutant la commande :
```shell
python script.py
```
3. Cliquez sur le bouton "Import Excel" pour sélectionner un fichier Excel contenant les données de couleur.
4. Les couleurs seront triées et affichées dans l'interface graphique.
5. Vous pouvez également choisir d'afficher les informations supplémentaires telles que le HUE, le RGB et le pourcentage de saturation en cliquant sur les boutons correspondants.
6. Les données triées peuvent être exportées au format CSV en cliquant sur le bouton "Exporter CSV".

## Remarque

Ce script nécessite l'installation de la bibliothèque Pandas version 1.3.3. Vous pouvez installer cette bibliothèque en exécutant `pip install pandas==1.3.3`.

## Suivez-nous

- [@breakingtechfr](https://twitter.com/BreakingTechFR) sur Twitter.
- [Facebook](https://www.facebook.com/BreakingTechFr/) likez notre page.
- [Instagram](https://www.instagram.com/breakingtechfr/) taguez nous sur vos publications !
- [Discord](https://discord.gg/VYNVBhk) pour parler avec nous !