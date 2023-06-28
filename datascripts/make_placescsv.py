#!/bin/env python3

from osgeo import ogr
import os
import csv

import settings


class PlacesIntersector:
    def run(self):
        print("PlacesIntersector")

        self.reproject(settings.INPUT_ZONESFILE, settings.REPROJECTED_ZONESFILE, settings.CTAZONES_SHAPEFILE_IDFIELD, settings.CTAZONES_SHAPEFILE_NAMEFIELD)
        self.reproject(settings.INPUT_CITYBOUNDS_SHP, settings.REPROJECTED_CITY_SHP, settings.CITYBOUNDS_IDFIELD, settings.CITYBOUNDS_NAMEFIELD)
        self.reproject(settings.INPUT_COUNTYBOUNDS_SHP, settings.REPROJECTED_COUNTY_SHP, settings.COUNTYBOUNDS_IDFIELD, settings.COUNTYBOUNDS_NAMEFIELD)

        self.findplaces(settings.REPROJECTED_CITY_SHP, settings.OUTPUT_CITYCSV, 'City')
        self.findplaces(settings.REPROJECTED_COUNTY_SHP, settings.OUTPUT_COUNTYCSV, 'County')

    def reproject(self, inputshp, outputshp, idfield, namefield):
        # reproject the shapefile to an Albers so we can do area calculations in findplaces()
        # and to standardize on there being only one attribute: name
        print("    Reproject {}  => {}".format(inputshp, outputshp))

        command = "{} {} -proj {} -filter-fields {} -rename-fields name={},id={} -o {} -quiet".format(
            settings.MAPSHAPER_CLI,
            inputshp,
            settings.PLANAR_SRS,
            ','.join([idfield, namefield]),
            namefield, idfield,
            outputshp
        )
        # print(command)
        os.system(command)

    def findplaces(self, placesdataset, csvfilename, placecolumnname):
        print("    Calculating {}  =>  {}".format(placesdataset, csvfilename))

        outfh = open(csvfilename, 'w')
        csvfh = csv.writer(outfh)
        csvfh.writerow(['ZoneIDOrig', placecolumnname])
        print(" settings.REPROJECTED_ZONESFILE", settings.REPROJECTED_ZONESFILE)
        # ctads = ogr.Open(settings.REPROJECTED_ZONESFILE, False)
        ctads = ogr.Open(settings.INPUT_ZONESFILE, False)

        ctalayer = ctads.GetLayer(0)

        for cta in ctalayer:
            ctaid = cta.GetField('id')
            ctageom = cta.GetGeometryRef()

            places = []

            ds = ogr.Open(placesdataset, False)
            layer = ds.GetLayer(0)
            layer.SetSpatialFilter(ctageom)

            for thisplace in layer:
                # work around twitchy hands making false intersections
                # "% of CTA area" strategy doesn't work: small towns in large rural CTAs = small percentage
                # but a town sliver over X acres, well, that should count as intersecting the town.
                #
                # also work around boundary datasets that are so precisely snapped,
                # that we get zero-area intersection as the overlapping boundary linestring of two areas
                # this leads to harmless but scary "non-surface geometry" warnings
                #
                # also, note that we collect names here and unique-ify them in a second step
                # multipolygon datasets means that a CTA may intersect the same place more than once!
                geom = thisplace.GetGeometryRef()
                intersection = geom.Intersection(ctageom)

                iacres = 0
                if intersection.GetGeometryName() in ('POLYGON', 'MULTIPOLYGON', 'GEOMETRYCOLLECTION'):
                    iacres = intersection.Area() * settings.SQMETERS_TO_ACRES

                if iacres < 2000:
                    continue

                name = thisplace.GetField('name')
                # print("            {}".format(name))
                places.append(name)

            ds = None  # close places dataset, will reopen at next CTA

            # done collecting: unique-ify the list, write the CSV rows
            places = list(set(places))
            for name in places:
                csvfh.writerow([ctaid, name])

        # done CTA loop, close geo fh and CSV fh
        ctads = None
        outfh.close()


if __name__ == '__main__':
    PlacesIntersector().run()
    print("DONE")
