# -*- coding: utf-8 -*-
import json
import pandas as pd
from shapely.geometry import shape, Polygon, MultiPolygon, mapping


def total_coords_count(geom):
    """Calcule le nombre total de points (coords) dans un Polygon."""
    if isinstance(geom, Polygon):
        rings = [geom.exterior] + list(geom.interiors)
        return sum(len(ring.coords) for ring in rings)
    return 0


def adaptive_polygon_simplify(geom, target_points=780, max_iterations=400):
    """
    Simplifie un Polygon jusqu'à atteindre environ `target_points` au total.
    """
    original = total_coords_count(geom)
    if original <= target_points:
        return geom, 0.0, original, original

    tolerance = 1e-10
    simplified = geom.simplify(tolerance, preserve_topology=True)
    iteration = 0

    while total_coords_count(simplified) > 780 and iteration < max_iterations:
        error_ratio = total_coords_count(simplified) / target_points
        tolerance *= min(error_ratio, 2)

        simplified = geom.simplify(tolerance, preserve_topology=True)
        iteration += 1

    simplified_n = total_coords_count(simplified)
    return simplified, tolerance, original, simplified_n


def geojson_to_excel_with_exploded_multipolygons(input_geojson_path, output_excel_path, output_geojson_path):
    with open(input_geojson_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Gérer une Feature unique ou une FeatureCollection
    if data.get("type") == "FeatureCollection":
        features = data["features"]
    elif data.get("type") == "Feature":
        features = [data]
    else:
        raise ValueError("Le fichier GeoJSON ne contient pas une Feature ou FeatureCollection valide.")

    all_records = []
    simplified_features = []

    for feature in features:
        props = feature.get("properties", {})
        geom = shape(feature["geometry"])

        # Éclatement des MultiPolygon en Polygon
        if isinstance(geom, MultiPolygon):
            polys = list(geom.geoms)
        elif isinstance(geom, Polygon):
            polys = [geom]
        else:
            continue  # ignorer les types non supportés

        for poly in polys:
            simplified_geom, tol, orig_pts, simp_pts = adaptive_polygon_simplify(poly)

            geom_json = mapping(simplified_geom)

            # Préparer la Feature simplifiée
            feature_geojson = {
                "type": "Feature",
                "geometry": geom_json,
                "properties": props
            }
            simplified_features.append(feature_geojson)

            # Préparer les données pour Excel
            record = props.copy()
            record_geojson = {
                "type": "Feature",
                "geometry": geom_json
            }
            record["geometry"] = json.dumps(record_geojson, ensure_ascii=False, separators=(',', ':'))
            if tol > 0:
                record["simplification_info"] = f"{orig_pts}→{simp_pts} points (tolérance={tol:.0e})"
            else:
                record["simplification_info"] = "Aucune simplification"

            all_records.append(record)

    # Export Excel
    df = pd.DataFrame(all_records)
    df.to_excel(output_excel_path, index=False)

    # Export GeoJSON
    final_geojson = {
        "type": "FeatureCollection",
        "features": simplified_features
    }
    with open(output_geojson_path, "w", encoding="utf-8") as f:
        json.dump(final_geojson, f, ensure_ascii=False, indent=2)


# Exemple d’appel à personnaliser :
geojson_to_excel_with_exploded_multipolygons(
    input_geojson_path=r"C:/Users/lgrillon/Downloads/SER__test.geojson",
    output_excel_path=r"C:/Users/lgrillon/Downloads/SER__test3.xlsx",
    output_geojson_path=r"C:/Users/lgrillon/Downloads/SER__test3.geojson"
)
