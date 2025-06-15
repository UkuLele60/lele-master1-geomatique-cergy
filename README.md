Ce code a été réalisé par chatgpt.
Il est adapté dans le cas où vous utilisez un fichier excel (ou un csv, mais dans ce cas il faut l'adapter un peu) comme database sur superset. 
Il permet : 
- de récupérer tous les champs existant d'un geojson et de créer les colonnes correspondantes sur excel
- de récupérer les géométries des objets présents dans le fichier geojson et les mettre dans le format spécial qui fonctionne pour le graphique deck.gl polygon sur Apache Superset, soit :
  {"type":"Feature","geometry":{"type":"Polygon","coordinates":[[[...]]]}}

