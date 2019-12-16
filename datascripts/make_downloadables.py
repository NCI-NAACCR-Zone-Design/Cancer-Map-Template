#!/bin/env python3

import os
import csv
import json
import zipfile
import copy

import settings


class DownloadableZipsMaker:
    def run(self):
        print("DownloadableZipsMaker")

        # self.statsbycta is our master stats storage from the CSVs, an assoc keying CTZ Zone ID => StatsDict
        # this is our source of truth for what zones should exist
        # all zones listed therein should have a presence in all CSV files, and all rows found in CSVs should have a presence in this assoc
        self.initializeCTAStatistics()

        # step by step...
        # note that this reads our own OUTPUT_ files, the very ones we're using in the website
        # as they have had their data massaged and standardized and some private data pruned
        # and well, it's what we want the public to have!
        self.aggregateDemographicData()
        self.aggregateCountyData()
        self.aggregateCityData()
        self.aggregateIncidenceData()
        self.createMasterShapefile()
        self.writeMasterZipFile()
        self.writeIndividualZipFiles()

    def initializeCTAStatistics(self):
        self.statsbycta = {}

        print("Loading geodata list of all CTAs")
        with open(settings.OUTPUT_CTATOPOJSON, 'r') as jsonfh:
            places = json.load(jsonfh)
            for place in places['objects']['ctazones']['geometries']:
                zoneid = place['properties']['Zone']
                self.statsbycta[zoneid] = {
                    'Zone': zoneid,
                    'URL': "{}?address={}".format(settings.WEBSITE_URL, zoneid),
                }

        # add placeholder for CA
        self.statsbycta['Statewide'] = {
            'Zone': 'Statewide',
            'URL': settings.WEBSITE_URL,
        }

    def aggregateIncidenceData(self):
        # add to self.statsbycta the incidence data for each one
        print("    Loading incidence CSV")
        with open(settings.OUTPUT_INCIDENCECSV, 'r') as csvfh:
            csvreader = csv.DictReader(csvfh)
            for row in csvreader:
                zoneid = row['Zone']
                if not zoneid:  # blank row
                    continue
                if zoneid not in self.statsbycta:
                    raise ValueError("Incidence CSV: CTA Zone ID {} not found in geodata".format(zoneid))

                if 'incidence' not in self.statsbycta[zoneid]:  # start the list if we've not seen this CTA yet
                    self.statsbycta[zoneid]['incidence'] = []

                # some data fixes, because they don't handle No Data consistently
                # sometimes the string "null" or "Null" or whatnot; sometimes just a blank cell
                for k, v in row.items():
                    if v == '' or v.strip().lower() == 'null':
                        row[k] = None

                # more data workarounds: some of the fields are integer/float so let's make sure they're kept that way
                thisdatarow = {
                    'Sex': row['sex'],
                    'Cancer': row['cancer'],
                    'Years': row['years'],
                    'PopTot': int(float(row['PopTot'])) if row['PopTot'] is not None else None,
                    'AAIR': float(row['AAIR']) if row['AAIR'] is not None else None,
                    'LCI': float(row['LCI']) if row['LCI'] is not None else None,
                    'UCI': float(row['UCI']) if row['UCI'] is not None else None,
                    'White_PopTot': int(float(row['W_PopTot'])) if row['W_PopTot'] is not None else None,
                    'White_AAIR': float(row['W_AAIR']) if row['W_AAIR'] is not None else None,
                    'White_LCI': float(row['W_LCI']) if row['W_LCI'] is not None else None,
                    'White_UCI': float(row['W_UCI']) if row['W_UCI'] is not None else None,
                    'Black_PopTot': int(float(row['B_PopTot'])) if row['B_PopTot'] is not None else None,
                    'Black_AAIR': float(row['B_AAIR']) if row['B_AAIR'] is not None else None,
                    'Black_LCI': float(row['B_LCI']) if row['B_LCI'] is not None else None,
                    'Black_UCI': float(row['B_UCI']) if row['B_UCI'] is not None else None,
                    'Hispanic_PopTot': int(float(row['H_PopTot'])) if row['H_PopTot'] is not None else None,
                    'Hispanic_AAIR': float(row['H_AAIR']) if row['H_AAIR'] is not None else None,
                    'Hispanic_LCI': float(row['H_LCI']) if row['H_LCI'] is not None else None,
                    'Hispanic_UCI': float(row['H_UCI']) if row['H_UCI'] is not None else None,
                    'Asian_PopTot': int(float(row['A_PopTot'])) if row['A_PopTot'] is not None else None,
                    'Asian_AAIR': float(row['A_AAIR']) if row['A_AAIR'] is not None else None,
                    'Asian_LCI': float(row['A_LCI']) if row['A_LCI'] is not None else None,
                    'Asian_UCI': float(row['A_UCI']) if row['A_UCI'] is not None else None,
                }

                self.statsbycta[zoneid]['incidence'].append(thisdatarow)

    def aggregateDemographicData(self):
        # add to self.statsbycta the demographic data for each one
        print("    Loading demographic CSV")
        with open(settings.OUTPUT_DEMOGCSV, 'r') as csvfh:
            csvreader = csv.DictReader(csvfh)
            for row in csvreader:
                zoneid = row['Zone']
                if zoneid not in self.statsbycta:
                    raise ValueError("Demographic CSV: CTA Zone ID {} not found in geodata".format(zoneid))

                self.statsbycta[zoneid]['demogs'] = {
                    'PopAll': int(row['PopAll']),
                    'QNSES': int(row['QNSES']) if row['QNSES'] != '' else None,
                    'PerRural': round(float(row['PerRural']), 1),
                    'PerUninsured': round(float(row['PerUninsured']), 1),
                    'PerForeignBorn': round(float(row['PerForeignBorn']), 1),
                    'PerWhite': round(float(row['PerWhite']), 1),
                    'PerBlack': round(float(row['PerBlack']), 1),
                    'PerHispanic': round(float(row['PerHispanic']), 1),
                    'PerAsian': round(float(row['PerAPI']), 1),
                }

    def aggregateCountyData(self):
        # add to self.statsbycta the counties each one intersects
        print("    Loading county CSV")
        with open(settings.OUTPUT_COUNTYCSV, 'r') as csvfh:
            csvreader = csv.DictReader(csvfh)
            for row in csvreader:
                zoneid = row['Zone']
                if zoneid not in self.statsbycta:
                    raise ValueError("County CSV: CTA Zone ID {} not found in geodata".format(zoneid))

                if 'counties' not in self.statsbycta[zoneid]:  # start the list if we've not seen this CTA yet
                    self.statsbycta[zoneid]['counties'] = []

                self.statsbycta[zoneid]['counties'].append(row['County'])

        # add stats for Statewide; we don't really load this but we need the placeholder row
        self.statsbycta['Statewide']['counties'] = []

    def aggregateCityData(self):
        # add to self.statsbycta the cities each one intersects
        print("Loading city CSV")
        with open(settings.OUTPUT_CITYCSV, 'r') as csvfh:
            csvreader = csv.DictReader(csvfh)
            for row in csvreader:
                zoneid = row['Zone']
                if zoneid not in self.statsbycta:
                    raise ValueError("City CSV: CTA Zone ID {} not found in geodata".format(zoneid))

                if 'cities' not in self.statsbycta[zoneid]:  # start the list if we've not seen this CTA yet
                    self.statsbycta[zoneid]['cities'] = []

                self.statsbycta[zoneid]['cities'].append(row['City'])

        # add stats for Statewide; we don't really load this but we need the placeholder row
        self.statsbycta['Statewide']['cities'] = []

    def createMasterShapefile(self):
        print("    Generating CTA shapefile")

        for ext in ['shp', 'shx', 'dbf', 'prj']:  # delete the target shapefile
            basename = os.path.splitext(settings.TEMP_CTASHPFILE)[0]
            if os.path.exists("{}.{}".format(basename, ext)):
                os.unlink("{}.{}".format(basename, ext))

        command = 'ogr2ogr -s_srs EPSG:4326 -t_srs EPSG:3310 -sql "SELECT Zone, ZoneName FROM {} ORDER BY Zone" {} {}'.format(
            'ctazones',
            settings.TEMP_CTASHPFILE,
            settings.OUTPUT_CTATOPOJSON
        )
        # print(command)
        os.system(command)

    def writeMasterZipFile(self):
        print("    Creating master ZIP")

        zipfilename = os.path.join(settings.DOWNLOADS_DIR, settings.MASTER_ZIPFILE_FILENAME)
        targetzip = zipfile.ZipFile(zipfilename, 'w', zipfile.ZIP_DEFLATED, 9)

        print("        Geodata")
        for ext in ['shp', 'shx', 'dbf', 'prj']:
            basename = os.path.splitext(settings.TEMP_CTASHPFILE)[0]
            realfilename = "{}.{}".format(basename, ext)
            inzipfilename = os.path.basename("{}.{}".format(basename, ext))
            targetzip.write(realfilename, inzipfilename)

        print("        CSV content")
        csvfilename = settings.MASTER_CSV_FILENAME
        zoneid = 'ALL'
        self.fetchDataAndSaveCsk(zoneid, csvfilename)
        targetzip.write(csvfilename, os.path.basename(csvfilename))

        print("        Metadata document")  # not so easy! explicitly convert to CRLF linefeeds for poor folks stil using Notepad
        with open(settings.DOWNLOADZIP_READMEFILE, 'r') as textfh:
            readmetext = textfh.read()
            readmetext = readmetext.replace('\n', '\r\n')
        targetzip.writestr(os.path.basename(settings.DOWNLOADZIP_READMEFILE), readmetext)

        targetzip.close()
        print("        Created {}".format(zipfilename))

    def writeIndividualZipFiles(self):
        print("    Creating individual CTA ZIPs")

        with open(settings.DOWNLOADZIP_READMEFILE, 'r') as textfh:
            readmetext = textfh.read()
            readmetext = readmetext.replace('\n', '\r\n')

        for zoneid in self.statsbycta.keys():
            if zoneid == 'Statewide':  # CA is in the Statewide/All file but not a single-zone file
                continue

            csvfilename = settings.PERCTA_CSV_FILENAME.format(zoneid)
            zipfilename = os.path.join(settings.DOWNLOADS_DIR, settings.PERCTA_ZIPFILES_FILENAME.format(zoneid))
            print("        {}".format(zipfilename))

            targetzip = zipfile.ZipFile(zipfilename, 'w', zipfile.ZIP_DEFLATED, 9)

            self.fetchDataAndSaveCsk(zoneid, csvfilename)
            targetzip.write(csvfilename, os.path.basename(csvfilename))

            targetzip.writestr(os.path.basename(settings.DOWNLOADZIP_READMEFILE), readmetext)

            targetzip.close()

    def fetchDataAndSaveCsk(self, zoneid, csvfilename):
        # the big wrapper to fetch CSV rows for a CTA zone, generate CSV-shaped content, and save to disk
        headrow = self.csvHeaderRow()
        datarows = self.fetchRowsForCTA(zoneid)
        datarows = self.formatRowsForCsv(datarows)
        self.writeCsvToDisk(csvfilename, headrow, datarows)

    def csvHeaderRow(self):
        # not coincidentally, these column headings are also key names for rows returned by fetchRowsForCTA()
        # meaning this array can be used as CSV header, and also looped over to fetch fields when generating a CSV
        # in fact, see formatRowsForCsv() where we do exactly that
        return [
            'Zone',
            'Counties',
            'Cities',
            'URL',
            'Sex',
            'Cancer',
            'Years',
            'PopTot',
            'AAIR',
            'LCI',
            'UCI',
            'White_PopTot',
            'White_AAIR',
            'White_LCI',
            'White_UCI',
            'Black_PopTot',
            'Black_AAIR',
            'Black_LCI',
            'Black_UCI',
            'Hispanic_PopTot',
            'Hispanic_AAIR',
            'Hispanic_LCI',
            'Hispanic_UCI',
            'Asian_PopTot',
            'Asian_AAIR',
            'Asian_LCI',
            'Asian_UCI',
            'QNSES',
            'PopAll',
            'PerRural',
            'PerUninsured',
            'PerForeignBorn',
            'PerWhite',
            'PerBlack',
            'PerAsian',
            'PerHispanic',
            'PerAsian',
        ]

    def fetchRowsForCTA(self, zoneid):
        # this is the wrapper that translates the jumble of incidence, demogs, counties, cities into a set of rows for a given CTA
        # output is a list of dicts, each dict being a row for the CSV for a combinations of Cancer and Sex
        # these row-dicts will include demographics, counties, cities which will be the same for each row, cuz they're constant for the CTA not by cancer

        # the special zoneid 'ALL' means to return all rows for all CTAs, e.g. for the master ZIP
        if zoneid == 'ALL':
            thezoneids = self.statsbycta.keys()
        else:
            thezoneids = [zoneid]

        # loop over the zone IDs we should be fetching; per above, usually one but maybe all
        collectedrows = []
        for thiszoneid in thezoneids:
            if thiszoneid not in self.statsbycta:
                raise ValueError("fetchRowsForCTA(): CTA {} not found".format(thiszoneid))

            thiszonestats = self.statsbycta[thiszoneid]
            for incidencerow in thiszonestats['incidence']:
                # start with the incidence data
                thisrow = copy.deepcopy(incidencerow)
                # add these constants
                thisrow['Zone'] = thiszonestats['Zone']
                thisrow['URL'] = thiszonestats['URL']
                # add cities and counties
                thisrow['Counties'] = ", ".join(thiszonestats['counties']) if 'counties' in thiszonestats else None
                thisrow['Cities'] = ", ".join(thiszonestats['cities']) if 'cities' in thiszonestats else None
                # add demographics
                for k, v in thiszonestats['demogs'].items():
                    thisrow[k] = v
                # done, append
                collectedrows.append(thisrow)

        return collectedrows

    def formatRowsForCsv(self, rows):
        whichfields = self.csvHeaderRow()

        # could do a [] within a [] but that turns into an eyesore
        # ... and makes it tougher to modify individual fields and exceptions as the spec changes and changes and changes
        collectedrows = []
        for thisrow in rows:
            newrow = [thisrow[f] for f in whichfields]
            collectedrows.append(newrow)

        return collectedrows

    def writeCsvToDisk(self, outfilename, headrow, datarows):
        with open(outfilename, 'w') as csvfile:
            csvwriter = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC)
            csvwriter.writerow(headrow)
            for thisrow in datarows:
                csvwriter.writerow(thisrow)


if __name__ == '__main__':
    DownloadableZipsMaker().run()
    print("DONE")
