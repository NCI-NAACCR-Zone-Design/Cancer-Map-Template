
# Cancer Mapping Template

This is a template for web developers to set up a website that displays cancer statistics by cancer reporting zone. Note that throughout the template, cancer reporting zones are often referred to as CTAs (Cancer Tabulation Areas).

This is not a turnkey product with a 5-minute installer. It is a starting place for a web developer to set up a cancer mapper and to begin customizing their own website.

See the project on Github at https://github.com/NCI-NAACCR-Zone-Design/Cancer-Map-Template/.

See a demonstration at https://nci-naaccr-zone-design.github.io/Cancer-Map-Template/.

This template was developed based on the California Health Maps website at https://www.californiahealthmaps.org/.


## Prerequisites

### Software Setup

You need the **NVM** and **Yarn** command-line tools installed. To check, run `yarn --version` and `nvm --version`

You will need Python 3 in order to run the data-preparation scripts under `datascripts/`. To check, run `python3 --version` and `pip3 -version`

You will need the OSGEO/GDAL module for Python 3. To check, run `python3 -c 'from osgeo import ogr; print("OK")'` If it is not installed on your system, you will need to install GDAL and then Python's GDAL package. To install GDAL, see https://trac.osgeo.org/gdal/wiki/DownloadingGdalBinaries for recommended packages for various operating systems, including Mac and Windows. To install the Python library, run `pip3 install GDAL`

#### A note about Yarn on Ubuntu and Windows Subsystem for Linux (WSL)

If you are using Ubuntu or WSL, there is a different and completely unrelated *yarn* command as part of the *cmdtest* package. _Do not use the Yarn that comes with cmdtest._ If your `yarn` says that `-i` is not a known option, you're running the wrong one and you need to install it as follows:

```
sudo apt remove cmdtest
sudo apt remove yarn
curl -sS https://dl.yarnpkg.com/debian/pubkey.gpg | sudo apt-key add -
echo "deb https://dl.yarnpkg.com/debian/ stable main" | sudo tee /etc/apt/sources.list.d/yarn.list
sudo apt update && sudo apt install --no-install-recommends yarn
```

### Data

You will need to provide a number of data tables and geographic boundary files to supply content for the website.  These are described briefly here -- see the *Integrating Your Own Data* section of this document for more details:

* A CSV file of cancer incidence statistics for each cancer reporting zone (CTA).  For these cancer incidence statistics, you may find it helpful to have a companion file that lists all domain values for the cancer sites, sexes, year ranges, and race/ethnicities.  

* A CSV file of demographic statistics for each zone.  For these demographics data, you may find it helpful to have a companion file that lists the demographic fields that are included, how you would like them labelled and formatted, and explanatory tooltip content.

* A shapefile describing the geographic boundaries of the cancer reporting zones (CTAs).

* A shapefile describing the geographic boundaries of the counties for your state.

* A shapefile describing the geographic boundaries of the cities/Census Designated Places (CDPs) for your state.

### Hosting

You will need a place to host the website.

The resulting files which will comprise the website, are static files and no database nor server-side scripting is required. As such, the website files may be hosted on any web server or many serverless systems such as Amazon S3.

An overview of some options are listed below. For more details, see the **Deployment** section.
* **Github Pages** You need to set up a Github repository where this will be hosted. The repository may be private. It must have Github Pages enabled and set to serve from the `docs/` directory (not the `gh-pages` branch). Github Pages is free to use, if your soure code repository is public.
* **Commodity Web Hosting** The resulting website is static HTML files, and can be hosted on any commodity web server such as  Dreamhost, Bluehost, or HostGator. Pricing varies, with many options as little as $5 per month, and web hosts provide technical support if you have trouble getting your website files online.
* **Amazon S3** Amazon S3 is a very low-price option for hosting static files, and it can be configured to serve your website. Though more involved to set up, hosting is very inexpensive, as little as $0.50 per month. You will need to sign up for Amazon Web Services and create a S3 bucket, then configure that bucket for website hosting. For more information, see https://docs.aws.amazon.com/AmazonS3/latest/dev/WebsiteHosting.html


## Getting Started

Visit https://github.com/GreenInfo-Network/Westat-Cancer-Template Download and unpack the latest release ZIP file.

Note: If you are using a Mac and you choose to move files out of that folder into some other folder, there are files starting with a `.` and these will not be visible in Mac's Finder by default. Use Command+Shift+Dot to show these files, or use the Console instead.

Open your command-line tools and `cd` into the directory.

Select the appropriate Node version: `nvm use`

Install dependencies: `yarn -i`

Start the Webpack development web server: `npm start` This will run a web server at http://localhost:8181/ where you can see your website under development.

The package comes with a working dataset to serve as a demonstration and example. After a brief overview of the code, your next step will be to integrate your own data.



## Overview of Directory Structure and Development

### Directory Structure

The `datascripts/` folder has scripts used when integrating your own data and updating it.

The `src/` directory contains the source code files, including the settings for what fields to expect in your demographic and incidence CSVs, the text/HTML that displays, and calculations and map controls. Later on in this document are some task-oriented tips on how you would make edits to your website.

The `static/data/` directory contains the data files that power the website: the CTA Zones geodata file, the incidence and demographics CSVs, the county reference overlay, and so on.

The `static/downloads/` directory, contains content accessed via the *Download* button on the website.

The rest of the `src/static/` directory is where you should put other static content such as images, logos, and your favicon.

### Development

Program code under `src/` is written in ECMAScript 2017 and SASS, and compiled using Babel, Webpack, et al.

The command `npm start` will start Webpack's development server at http://localhost:8181/ and will open a browser window for you as well.

The command `npm run build` will copy static assets and compiled code into the `docs/` directory, where it may be deployed to Github Pages as your website. You will need to do this every time you change the content of `static/`, including replacing images or loading new data.

While using the Webpack development server, making edits to `src/` will automatically recompile your application and reload the web page. **However, changes to `static/` will not trigger this**. After replacing a logo, for example, you must run `npm run build` and reload the page.

Again, **do not forget to do `npm run build`** after making changes the content of `static/`, including replacing images or loading new data.

### Deployment

The command `npm run build` will compile the source files and static assets into their browser-ready versions under `docs/`. The contents of this folder are the ready-to-run website; no server-side database nor scripting services are required.

* **Github Pages** After using `npm run build`, run `git commit` and `git push` as usual, and Github Pages will update your website within 5 minutes. The command `npm run deploy` is a convenient shortcut: it will run the commands to build, add, commit, and push in one single command.
* **Commodity Web Hosting** The contents of this folder (not the folder itself) may be uploaded into your hosting directory via FTP/SFTP using a file transfer client such as WinSCP or FileZilla.
* **Amazon S3** The contents of this folder (not the folder itself) may be uploaded into your S3 bucket, via the S3 console or via a graphical client such as Cyberduck, S3 Browser, or DragonDisk.
 

## Integrating Your Own Data

Again, **do not forget to do `npm run build`** after making changes the content of `static/`, including replacing images or loading new data.

### Incidence Data

Cancer incidence rates are provided in a CSV file.  Zone-level rates can be calculated by aggregating tract-level data using the `ZonedTracts` crosswalk file developed during the cancer reporting zone definition process.  The cancer incidence file has the following field and data requirements:

* The `Zone` field is a text string used as the CTA Zones' unique ID to tie the incidence data to other data (demographic data, zone boundaries, etc.).

* The special `Zone` name *Statewide* should be used to indicate statewide data and the special `Zone` name *Nationwide* should be used to indicate nationwide data. 

* The `Cancer` field is a text string specifying domain values for the cancer site and is used for filtering.

* The `Sex` field is a text string specifying domain values for sex (female, male, and both sexes combined) and is used for filtering.

* The `Years` field is a text string specifying domain values for the range of years and is used for filtering.

* The incidence data fields `Cases` (number of cancer cases), `AAIR` (age-adjusted incidence rate per 100,000), `LCI` (95% lower confidence interval), `UCI` (95% upper confidence interval), and `PopTot` (population denominator) are numeric values used for reporting incidence. 

* The same incidence fields must be defined for each race/ethnicity filter that you will define, and must be prefixed by the race/ethnicity's "short version". For example, If you use `W` as a domain value for Non-Hispanic Whites then cancer rates for Non-Hispanic Whites will be reported using these fields: `W_Cases`, `W_AAIR`, `W_LCI`, `W_UCI`, and `W_PopTot`.

Copy your cancer incidence CSV into `static/data/cancerincidence.csv`

Edit `index.js` and set up `SEARCHOPTIONS_CANCERSITE` to match your dataset's domain values.  This variable contains a row for each cancer site in which the `value` field gives the short name for the cancer site used in the data file and the `label` field gives a longer cancer site name used for filtering and displaying the data.

Edit `index.js` and set up `SEARCHOPTIONS_SEX` to match your dataset's domain values. The `value` field gives the short text string used in the data file and the `label` field gives a longer text string used for filtering and displaying the data.

If any of the cancer site options will be specific to one sex, edit `index.js` and set up `CANCER_SEXES` to auto-select that sex if that cancer site is selected. 

Edit `index.js` and set up `SEARCHOPTIONS_TIME` to match your dataset's year range values. The set of year ranges must match between the incidence and demographic datasets. The `value` field gives the short text string used in the data file and the `label` field gives a longer text string used for filtering and displaying the data.

Edit `index.js` and set up `SEARCHOPTIONS_RACE` to match your dataset's domain values. The `value` field gives the short text string (usually a single letter) used as a race/ethnicity prefix for the five cancer incidence fields in the data file and the `label` field gives a longer race/ethnicity description used for filtering and displaying the data.  A null value (‘’) is usually used for all race/ethnicities.  

Edit `index.js` and set `NATIONWIDE_INCIDENCE` to indicate whether your data will support Nationwide readouts for comparison with the Zone and State statistics.

Again, **do not forget to do `npm run build`** after making changes to the content of `static/`, including replacing images or loading new data.

### Demographic Data

Demographic and socioeconomic data to be displayed in a table below the cancer incidence data are provided in a CSV file.  Zone-level data can be calculated by aggregating tract-level data using the ZonedTracts crosswalk file developed during the cancer reporting zone definition process.  The demographic data file has the following field and data requirements:

* The `Zone` field is a text string used as the CTA Zones' unique ID to tie the demographic data to other data (incidence data, zone boundaries, etc.).

* The special `Zone` name *Statewide* should be used to indicate statewide data. The special `Zone` name *Nationwide* should be used to indicate nationwide data. 

* The `Years` field is a text string specifying domain values for the range of years and is used for filtering. Values must match those used for the incidence data.  

* A set of numeric fields specify the values for each demographic data item.  The data items are defined in the `DEMOGRAPHIC_TABLES` as described below.

Copy your demographics CSV into `static/data/demographics.csv`

Edit `index.js` and set up `DEMOGRAPHIC_TABLES` to define the demographic data items to be displayed. This variable consists of a set of groups of demographic data items with a `title` field that gives a group name for a set.  Each group contains a set of rows of data items in which the `field` field gives the short name for the data item used in the data file, the `label` field gives a longer description of the data item used for filtering and displaying the data, the `format` field specifies the display format, and the `tooltip_id` field points to tooltip text in `index.html`. 

Edit `index.js` and set up `CHOROPLETH_OPTIONS` to include a row for each demographic data item with `field`, `label`, and `format` fields matching those in the `DEMOGRAPHIC_TABLES` variable and an additional `colorramp` field that specifies the `Color By` option to use when displaying the data item in the choropleth map.

Edit `index.js` and make sure `SEARCHOPTIONS_TIME` values match your dataset's domain values. The set of dates must match between the incidence and demographic datasets.

Edit `index.js` and set `NATIONWIDE_DEMOGRAPHICS` to indicate whether your data will support Nationwide readouts for comparison with the Zone and State statistics.

Edit `index.html` to specify tooltip text for each demographic measure.  The text should include a description of the measure and information about the data source.

Again, **do not forget to do `npm run build`** after making changes to the content of `static/`, including replacing images or loading new data.

### CTA Zones Geodata

Place your CTA Zones shapefile describing the boundaries of the cancer reporting zones into `datascripts/inputs/` as `CTAZones.shp`.

This should be provided in WGS84 (plain latitude-longitude / unprojected) spatial reference system (SRS).

Relevant attributes are as follows. Other fields will be ignored.

* `Zone` -- CTA Zone's unique ID, used to tie to other data (incidence, demographics).

* `ZoneName` -- CTA Zone's name for display.

Run `python3 make_ctageofile.py`. This will create `static/data/cta.json` which is the TopoJSON file providing CTA Zone boundaries for the map.

Again, **do not forget to do `npm run build`** after making changes to the content of `static/`, including replacing images or loading new data.

### County Boundaries Geodata

Place your county boundaries shapefile into `datascripts/inputs/` as `counties.shp`.

A county boundaries shapefile was probably provided with the final results of the zone definition process.  Alternatively, one can be obtained from ftp://ftp2.census.gov/geo/tiger/TIGER2019/COUNTY/ The FTP site has one county file for all of the United States, and you will need to crop it to your state using the `STATEFP` field.

This should be provided in WGS84 (plain latitude-longitude / unprojected) SRS.

Relevant attributes are as follows. Other fields will be ignored.

* `COUNTYFP` -- The FIPS code for this county. Used as a unique ID.

* `NAME` -- The name of the county.

Run `python3 make_countygeofile.py` to create `static/data/countybounds.json` which is the TopoJSON file providing county boundaries for the map.

Again, **do not forget to do `npm run build`** after making changes to the content of `static/`, including replacing images or loading new data.

### City / Place Boundaries Geodata

Place your city/CDP boundaries shapefile into `datascripts/inputs/` as `cities.shp`.

A city/CDP boundaries shapefile was probably provided with the final results of the zone definition process.  Alternatively, one can be obtained from ftp://ftp2.census.gov/geo/tiger/TIGER2019/PLACE/ The FTP site names the files by the state's FIPS code, e.g. California is FIPS code `06`.

This should be provided in WGS84 (plain latitude-longitude / unprojected) SRS.

Relevant attributes are as follows. Other fields will be ignored.

* `PLACEFP` -- The FIPS code for this county. Used as a unique ID.

* `NAME` -- The name of the city/place.

After the CTA zones, counties, and places shapefiles are in place, run `python3 make_placescsv.py` to create `static/data/counties_by_cta.csv` and `static/data/cities_by_cta.csv` which provide a list of places intersecting each CTA Zone.  Note that, prior to determining the intersections, this script sets a projection for each shapefile using the projection specified in `PLANAR_SRS` in the `settings.py` script.  The default projection is Texas Centric Albers Equal Area (https://spatialreference.org/ref/epsg/3083/).  Other Albers equal area projections could be specified.  

Again, **do not forget to do `npm run build`** after making changes to the content of `static/`, including replacing images or loading new data.

### Creating Downloadable Files

The files offered by the Download button, are static ZIP files containing CSV extracts of merged demographic and incidence data.

Edit `datascripts/make_downloadables.py` and define what demographic and incidence fields should be present in the downloadable content accessed by the Download button.
* The function `aggregateDemographicData()` will read the demographic dataset and is where you can massage/correct/format the data for the downloadable CSVs, as well as rename the fields as they appear in the download CSVs.
* The function `aggregateIncidenceData()` does simialrly, for the incidence data, massaging and formatting values and renaming them for the download CSVs.
* The function `csvHeaderRow()` defines the sequence of fields as they appear in the final download CSV. All fields here must be the fields created in `aggregateIncidenceData()` and/or `aggregateDemographicData()` but it is not required that every field defined be used here.

Edit the `datascripts/inputs/readme.txt` file to describe the CSV fields, and to include the name of your website or project as needed. This file will be included in all of the downloadable ZIP files, and is suitable for metadata such as a data dictionary, a disclaimer, and credits/attributions.

Run `python3 make_downloadables.py` to compile the downloadable ZIP files under `static/downloads/`.

Again, **do not forget to do `npm run build`** after making changes the content of `static/`, including replacing images or loading new data.

### Rebuilding For The Website

Lastly, be sure to run `npm run build` to update the files as seen by the web server.

Again, **do not forget to do `npm run build`** after making changes the content of `static/`, including replacing images or loading new data.



## Further Customizations

### HTML Site/Registry Name and Statistics

The file `src/index.html` contains a boilerplate version of site copy, with several places where you will want to enter the name of your website/registry/project, statistical thresholds and values, and so on.

Search for the `[REPLACE` string throughout `src/index.html` and replace the values mentioned there. Some are simple values amenable to simple search-and-replace, but most are narrative text that may require more involvement.

A list of such replacements is:
* `[REPLACE STATE/REGISTRY]` -- The name of your state, project, or cancer registry. Commonly used with the phrase "Cancer Maps" after it, indicating the name of this website. A simple search-and-replace should work here.
* `[REPLACE NUM_CANCER_SITES]` -- The number of cancer sites by which data may be searched. Usually the same as the number of `SEARCHOPTIONS_CANCERSITE` entries.
* `[CONFIRM RACE LIST]` -- A list of the races/ethnicities by which data may be searched. This should reflect the `SEARCHOPTIONS_RACE` entries.
* `[REPLACE NUM_ZONES]` -- The number of CTA Zones used for the analysis.
* `[REPLACE MINIMUM ZONE POPULATION]` -- The minimum population of a CTA Zone. Used in a statement describing CTA Zones.
* `[REPLACE MAXIMUM ZONE POPULATION]` -- The maximum population of a CTA Zone. Used in a statement describing CTA Zones.
* `[REPLACE MINIMUM TRACTS PER ZONE]` -- The minimum number of census tracts forming any CTA Zone. Used in a statement describing CTA Zones.
* `[REPLACE MAXIMUM TRACTS PER ZONE]` -- The maximum number of census tracts forming any CTA Zone. Used in a statement describing CTA Zones.
* `[REPLACE REPORTING MIN CASES]` -- The minimum number of cancer cases in a CTA Zone, to be reported.
* `[REPLACE REGISTRY WEBSITE]` -- A hyperlink URL to this website's parent agency, cancer registry, etc., in the FAQ.
* `[REPLACE FUNDING URL]` -- A hyperlink URL to the agency which funded this website. Displayed in the FAQ alongside the FUNDING SOURCE.
* `[REPLACE FUNDING SOURCE]` -- A statement/description of who funded the website. Displayed in the FAQ alongside the FUNDING URL.
* `[REPLACE ABOUT BLURB]` -- A statement/description of the website, in "What is the XXX Registry" section of the FAQ.
* `[REPLACE CITATION INFO]` -- A statement/description of how this website should be cited in literature.
* `[REPLACE NATIONAL CANCER DATA SOURCE INFO]` -- A statement/description of the national cancer data source, including data years. This should be reviewed with subsequent data updates to verify whether it needs to be updated as well (e.g., during annual data updates). An example of this statement applicable to national cancer data through 2018 is: "National incidence data come from the National Program of Cancer Registries and Surveillance, Epidemiology, and End Results SEER*Stat Database: U.S. Cancer Statistics Incidence Analytic file - 1998-2018. United States Department of Health and Human Services, Centers for Disease Control and Prevention. Released June 2021, based on the 2020 submission."

Since this is free-form narrative text, you may choose to rewrite or rephrase whole blocks of text, in addition to or instead of making the string replacements suggested above.


### Other HTML Content, Cosmetic Changes, Look-and-Feel

* *Browser title bar* -- Look in `src/index.html` for the `title`.

* *Footer, credits, and citation* -- Look in `src/index.html` for the `footer`.

* *Favicon* -- Replace or change the refgerence to `/static/favicon.png` with an appropriate image. Don't forget to `npm run build`. The `/static` directory includes several boilerplate logos if desired.

* *Introductory text/logo/navbar* -- Look in `src/index.html` for the `intro-text` section. The `/static` directory includes several boilerplate logos if desired.

* *Map starting view* -- Look in `src/index.js` for the definition of `MAP_BBOX` which defines lat-lng coordinates for `[[south, west], [north, east]]` The website http://bboxfinder.com is very useful here. *Note that the actual bounding box viewed depends on a lot of factors such as the size of the browser window, so the map view may not be precisely what you want and may not be the same on different displays.*

* *Google Analytics* -- Look in `src/index.html` for a `script` tag pointing at *www.googletagmanager.com* Fill in your UA code _in two places_ here.

* *Bing API Key* -- Look in `src/index.html` for the definition of `BING_API_KEY` Until you set this, you will not be able to search for addresses. A Bing Maps API key is free, and their terms of use are quite flexible. See https://docs.microsoft.com/en-us/bingmaps/getting-started/bing-maps-dev-center-help/getting-a-bing-maps-key for more information.

* *About This Project* -- Look in `src/index.html` for the `learn-about`.

* *Methodology* -- Look in `src/index.html` for the `learn-method`.

* *FAQs* -- Look in `src/index.html` for the `learn-faq`.

* *Glossary* -- Look in `src/index.html` for the `learn-glossary`.

* *Tooltip i icons* -- Within `src/index.html` you may create tooltip I icons, with HTML such as this: `<i class="fa fa-info-circle" aria-hidden="true" data-tooltip="yourtermhere"></i>` The tooltip HTML for each such tooltip, is provided in `tooltip_contents` Each DIV has a `data-tooltip` attribute corresponding to the `data-tooltip` used in the `<i>` element. For the Demographics table, be sure to cross reference to the `DEMOGRAPHICS_TABLE` element in `src/index.js`.

### Downloadable Data Files

* *readme.txt* -- See `datascripts/readme.txt`
