#!/bin/env python

# the destination CSV and TopoJSON files which we will generate
OUTPUT_CTATOPOJSON = "../static/data/cta.json"
OUTPUT_INCIDENCECSV = "../static/data/cancerincidence.csv"
OUTPUT_DEMOGCSV = "../static/data/demographics.csv"
OUTPUT_COUNTYCSV = "../static/data/counties_by_cta.csv"
OUTPUT_CITYCSV = "../static/data/cities_by_cta.csv"
OUTPUT_COUNTYJSON = "../static/data/countybounds.json"

# the downloadable ZIP files; one of all data, and for each CTA Zone
# all ZIP files will have a readme file added for credits, disclaimer, metadata
DOWNLOADS_DIR = "../static/downloads"
DOWNLOADZIP_READMEFILE = "./inputs/readme.txt"

MASTER_ZIPFILE_FILENAME = "all_zones.zip"
MASTER_CSV_FILENAME = "./tempfiles/all_zones.csv"  # include the ./tempfiles/ path; it will be stripped in the ZIP

PERCTA_ZIPFILES_FILENAME = "zone_{}.zip"
PERCTA_CSV_FILENAME = "./tempfiles/zone_{}.csv"  # include the ./tempfiles/ path; it will be stripped in the ZIP

TEMP_CTASHPFILE = './tempfiles/CTAZones_Download.shp'

# in the downloaded ZIP files, a CSV field will be a URL
# to link back to this website zoomed to a CTA Zone
WEBSITE_URL = "https://www.example.com"

# CTA Zones shapefile, and which fields to use from it
INPUT_ZONESFILE = "./inputs/CTAZones.shp"
CTAZONES_SHAPEFILE_IDFIELD = "ZoneIDOrig"
CTAZONES_SHAPEFILE_NAMEFIELD = "ZoneName"
REPROJECTED_ZONESFILE = "./tempfiles/ctazones2.shp"

# Census Designated Places shapefile and the County shapefile
# used for the CTA-to-City and CTA-to-County CSV lookup CSVs
INPUT_COUNTYBOUNDS_SHP = "./inputs/counties.shp"
COUNTYBOUNDS_IDFIELD = "COUNTYFP"
COUNTYBOUNDS_NAMEFIELD = "NAME"
REPROJECTED_COUNTY_SHP = "./tempfiles/counties2.shp"

INPUT_CITYBOUNDS_SHP = "./inputs/cities.shp"
CITYBOUNDS_IDFIELD = "PLACEFP"
CITYBOUNDS_NAMEFIELD = "NAME"
REPROJECTED_CITY_SHP = "./tempfiles/cities2.shp"

# TopoJSON settings for simplifying, quantizing coordinates, and rounding coordiate decimals
SIMPLIFY = "20%"
QUANTIZE = "1e5"
LATLNGPRECISION = 0.0001

# the path to the mapshaper CLI tool
# this should be in node_modules/.bin as it was installed via yarn/npm
MAPSHAPER_CLI = "../node_modules/.bin/mapshaper"

# choose a planar SRS which preserves area well, e.g. EPSG:3083 for Texas-centric AEA, or 3310 for California Teale-Albers
# this is used for finding the area of intersection, for finding city/county overlaps to CTA Zones
# PLANAR_SRS = "EPSG:3083"
PLANAR_SRS = "EPSG:2235"

# and when reprojecting to a globe (lat-long) the SRS to use
GLOBE_SRS = "EPSG:4326"

# other constants and calculations
SQMETERS_TO_ACRES = 0.000247105
