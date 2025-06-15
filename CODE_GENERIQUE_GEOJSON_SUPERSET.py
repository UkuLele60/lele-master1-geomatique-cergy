# -*- coding: utf-8 -*-
import json
import pandas as pd
from shapely.geometry import shape, Polygon, MultiPolygon, mapping

# ------------------------------------------
# Fonction pour compter le nombre total de points (couples de coordonnées)
# dans tous les anneaux (exterieur + intérieurs) d’un polygone. En effet, 
#dans l'arborescence d'un geojson, il y a "coordinates", qui mène soit 
#vers une liste de couple de coordonnées directement, soit vers des anneaux, 
#lorsque par exemple un polygone contient un trou à l'intérieur
# ------------------------------------------
def total_coords_count(geom):
    if isinstance(geom, Polygon):
        rings = [geom.exterior] + list(geom.interiors)
        return sum(len(ring.coords) for ring in rings)
    return 0  # Par défaut si ce n’est pas un Polygon

# ------------------------------------------
# Fonction de simplification adaptative :
# Réduit le nombre de points d’un polygone jusqu’à atteindre environ `target_points`
#Le target point est le nombre de point (couple de coordonnées) d'un polygone au delà duquel
#la géométrie est généralement trop grande pour être contenu dans une seule case d'un fichier excel,
#il faut donc simplifier et réduire le nombre de points.
# ------------------------------------------
def adaptive_polygon_simplify(geom, target_points=780, max_iterations=400):
    original = total_coords_count(geom)
    
    # Si déjà en dessous du seuil, ne rien faire
    if original <= target_points:
        return geom, 0.0, original, original

    tolerance = 1e-10  # tolérance initiale très faible
    simplified = geom.simplify(tolerance, preserve_topology=True)
    iteration = 0

    # On augmente progressivement la tolérance jusqu’à descendre sous 780 points
    while total_coords_count(simplified) > 780 and iteration < max_iterations:
        error_ratio = total_coords_count(simplified) / target_points
        tolerance *= min(error_ratio, 2)  # multiplier la tolérance doucement (x2 max)
        simplified = geom.simplify(tolerance, preserve_topology=True)
        iteration += 1

    simplified_n = total_coords_count(simplified)
    return simplified, tolerance, original, simplified_n

# ------------------------------------------
# Fonction principale : transforme un GeoJSON en Excel + nouveau GeoJSON simplifié
# ------------------------------------------
def geojson_to_excel_with_exploded_multipolygons(input_geojson_path, output_excel_path, output_geojson_path):
    # Chargement du fichier GeoJSON
    with open(input_geojson_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # On accepte soit une FeatureCollection, soit une Feature seule (quand le fichier
    #geojson contient un seul objet)
    if data.get("type") == "FeatureCollection":
        features = data["features"]
    elif data.get("type") == "Feature":
        features = [data]
    else:
        raise ValueError("Le fichier GeoJSON ne contient pas une Feature ou FeatureCollection valide.")

    all_records = []           # Contiendra les lignes pour Excel
    simplified_features = []   # Contiendra les features pour le GeoJSON simplifié

    for feature in features:
        props = feature.get("properties", {})  # On garde toutes les propriétés
        geom = shape(feature["geometry"])      # Conversion GeoJSON → objet shapely

        # Si MultiPolygon → on éclate en plusieurs Polygons, car superset ne gère pas 
        #les multipolygon
        
        if isinstance(geom, MultiPolygon):
            polys = list(geom.geoms)
        elif isinstance(geom, Polygon):
            polys = [geom]
        else:
            continue  # On ignore les types non gérés (Point, LineString, etc.)

        for poly in polys:
            # On applique la simplification adaptative
            simplified_geom, tol, orig_pts, simp_pts = adaptive_polygon_simplify(poly)

            # On convertit le polygone simplifié en GeoJSON natif
            geom_json = mapping(simplified_geom)

            # Préparer la feature simplifiée pour le GeoJSON final
            feature_geojson = {
                "type": "Feature",
                "geometry": geom_json,
                "properties": props
            }
            simplified_features.append(feature_geojson)

            # Préparer la ligne pour le fichier Excel
            record = props.copy()  # copier les propriétés originales
            record_geojson = {
                "type": "Feature",
                "geometry": geom_json
            }
            # Encodage JSON compressé pour l’intégration dans une cellule Excel
            record["geometry"] = json.dumps(record_geojson, ensure_ascii=False, separators=(',', ':'))

            # Ajouter une colonne d'information sur la simplification appliquée
            if tol > 0:
                record["simplification_info"] = f"{orig_pts}→{simp_pts} points (tolérance={tol:.0e})"
            else:
                record["simplification_info"] = "Aucune simplification"

            all_records.append(record)

    # Écriture du fichier Excel avec Pandas
    df = pd.DataFrame(all_records)
    df.to_excel(output_excel_path, index=False)

    # Création du nouveau GeoJSON simplifié
    final_geojson = {
        "type": "FeatureCollection",
        "features": simplified_features
    }
    with open(output_geojson_path, "w", encoding="utf-8") as f:
        json.dump(final_geojson, f, ensure_ascii=False, indent=2)

# ------------------------------------------
# chemins des fichiers (1 fichier en entrée et 2 fichiers en sortie)
# ------------------------------------------
geojson_to_excel_with_exploded_multipolygons(
    input_geojson_path=r"C:/Users/lgrillon/Downloads/SER__test.geojson",
    output_excel_path=r"C:/Users/lgrillon/Downloads/SER__test3.xlsx",
    output_geojson_path=r"C:/Users/lgrillon/Downloads/SER__test3.geojson"
)
