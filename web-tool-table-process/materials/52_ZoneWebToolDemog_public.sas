/* 52_ZoneWebToolDemog - Generate a table of demographic data for the web tool */

/* 5/25/2022 - added a Zone Design AUTOCALL macro library for commonly used macros */
/* 5/25/2022 - modified Excel file imports to use new ZD_ImportExcel macro */
/* 11/23/2022 - modified to use the web tool ZoneID without the StFIPS or the trailing repetition letters */

options sysprintfont=("Courier New" 8) leftmargin=0.75in nocenter compress=no;
ods graphics on;

%let stateName=Iowa;  /* Set state/registry name */
%let stateAbbr=IA;    /* Set state/registry abbreviation */
%let runNum=IA01; /* Run number used for Step 2 AZTool execution */
%let nationwide=yes;  /* Include nationwide data (yes|no)? */
/* Census 2010 variables to include: */
%let keepCens=PctRural PctMinority PctHispanic PctBlackNH;
/* ACS data and variables to include: */
%let oneACSPeriod=no;  /* Use one ACS period for all time periods (yes|no)? */
%let ACS_endyr1=2016;  /* ACS 5yr data period 1 end year (2016 = 2012-2016) */
%let ACS_endyr2=2011;  /* ACS 5yr data period 2 end year (2011 = 2007-2011) */
%let keepACS=Pct_forborn Pct100Pov PctNoHealthIns PctEducLHS;
%let keepACS2=Pct_forborn Pct100Pov PctEducLHS; /* PctNoHealthIns not available for 2011 */

/* Specify data path here for portability: */
%let pathbase=C:\Work\WebToolTables;

libname ZONEDATA "&pathbase.";
libname CENDATA1 "&pathbase.\Census_Data_Tables\ACS&ACS_endyr1._5yr Data";
%if &oneACSPeriod. = %quote(no) %then %do;
libname CENDATA2 "&pathbase.\Census_Data_Tables\ACS&ACS_endyr2._5yr Data";
%end; /* &oneACSPeriod=no processing */
ods pdf file="&pathbase.\52_ZoneWebToolDemog_&stateAbbr.to&ACS_endyr1..pdf";

/* Set up the Zone Design AUTOCALL macro library */
filename ZDAUTOS "&pathbase.\ZDMacros";
options mautosource sasautos=(SASAUTOS ZDAUTOS);

/* Import the final zoned tracts Excel file */
%ZD_ImportExcel(sourcefile=&pathbase.\ZonedTracts_&runNum._final.xlsx,
    sheetname=Tracts,
    targetds=ZonedTracts);

/* Import the final zone list Excel file */
%ZD_ImportExcel(sourcefile=&pathbase.\ZoneList_&runNum._final.xlsx,
    sheetname=FullStats,
    targetds=ZoneList);

/* Get ACS 5-year period 1 (most recent) datasets */
data ACSData_per1_Tract;
    set CENDATA1.ACS&ACS_endyr1._Tract;
run;
data ACSData_per1_State;
    set CENDATA1.ACS&ACS_endyr1._State;
run;
%if &nationwide. = %quote(yes) %then %do;
data ACSData_per1_Nation;
    set CENDATA1.ACS&ACS_endyr1._Nation;
run;
%end; /* &nationwide=yes processing */

%if &oneACSPeriod. = %quote(no) %then %do;
/* Get ACS 5-year period 2 (least recent) datasets */
data ACSData_per2_Tract;
    set CENDATA2.ACS&ACS_endyr2._Tract;
run;
data ACSData_per2_State;
    set CENDATA2.ACS&ACS_endyr2._State;
run;
%end; /* &oneACSPeriod=no processing */
%if &nationwide. = %quote(yes) & &oneACSPeriod. = %quote(no) %then %do;
data ACSData_per2_Nation;
    set CENDATA2.ACS&ACS_endyr2._Nation;
run;
%end; /* &nationwide=yes & &oneACSPeriod=no processing */

%if &nationwide. = %quote(yes) %then %do;
/* Import the US total demographic data table */
%ZD_ImportExcel(sourcefile=&pathbase.\Census_Data_Tables\DemogTable_US_Totals.xlsx,
    sheetname=DemogTable,
    targetds=DemogTable_US);
%end; /* &nationwide=yes processing */

/* Import WebTool table for years */
%ZD_ImportExcel(sourcefile=&pathbase.\CancerSiteTable_&stateAbbr..xlsx,
    sheetname=Webtool TIME,
    targetds=WebTool_Years);


/* Clean up the zone list dataset and keep just what we need */
data ZoneList2;
    set ZoneList;
	ZoneIDOrig = compress(ZoneIDOrig,'abcdefghijklmnopqrstuvwxyz'); /* 11/23/2022: Remove repetition letters */
    PctRural = 100 - PctUrban_PopMean;
    PctMinority = PctMinority_PopMean;
    PctHispanic = PctHispanic_PopMean;
    PctBlackNH = PctBlackNH_PopMean;
    PctAPINH = PctAPINH_PopMean;
    format PctRural PctMinority PctHispanic PctBlackNH PctAPINH 6.2;
    label
        PctRural = 'Percent of population living in a rural area'
        PctMinority = 'Percent minority (other than non-Hispanic White)'
        PctHispanic = 'Percent Hispanic'
        PctBlackNH = 'Percent non-Hispanic Black'
        PctAPINH = 'Percent non-Hispanic Asian/Pacific Islander';
    keep /* ZoneName */ ZoneIDOrig ZonePop
        &keepCens.; /* Just keep the variables needed for the web tool */
    rename ZoneIDOrig = ZoneID
        ZonePop = TotalPop;
run;

/* Clean up the zoned tracts dataset and keep just what we need */
data ZonedTracts2;
    set ZonedTracts;
	ZoneIDOrig = compress(ZoneIDOrig,'abcdefghijklmnopqrstuvwxyz'); /* 11/23/2022: Remove repetition letters */
    keep TractID ZoneIDOrig; /*11/23/2022: Updated to reference ZoneIDOrig variable*/
    rename ZoneIDOrig = ZoneID;
run;

/* Get the state FIPS code from the ACS state dataset and put it in a macro variable */
data _null_;
    set ACSData_per1_State;
    if GeoName = "&stateName.";
    putlog "*** &stateName. FIPS code set to " StFIPS;
    call symputx("stateFIPS", StFIPS);
run;

/* Macro to clean up the ACS datasets and keep just what we need */
%MACRO CleanACS(level=,period=);
data ACSData_per&period._&level.2;
    length ZoneID $10;
    set ACSData_per&period._&level.;
    if StFIPS in ("&stateFIPS.","00"); /* Just keep data for this state and the US */
    /* Just keep the variables needed for the web tool */
    %if &period. = %quote(1) %then %do;
        keep &keepACS.; /* Period 1 keeps */
        format &keepACS. 6.2;
    %end;
    %else %do;
        keep &keepACS2.; /* Period 2 keeps */
        format &keepACS2. 6.2;
    %end;
    %if &level. = %quote(Tract) %then %do;
        keep GeoID Total_pop;
        rename GeoID = TractID
            Total_pop = TotPop;
    %end;
    %else %do;
        ZoneID = "&level.wide";
        keep ZoneID;
    %end;
run;
%MEND CleanACS;

/* Clean up the period 1 ACS dataset and keep just what we need */
%CleanACS(level=Tract,period=1);
%CleanACS(level=State,period=1);
%if &nationwide. = %quote(yes) %then %do;
%CleanACS(level=Nation,period=1);
%end; /* &nationwide=yes processing */

%if &oneACSPeriod. = %quote(no) %then %do;
/* Clean up the period 2 ACS dataset and keep just what we need */
%CleanACS(level=Tract,period=2);
%CleanACS(level=State,period=2);
%end; /* &oneACSPeriod=no processing */
%if &nationwide. = %quote(yes) & &oneACSPeriod. = %quote(no) %then %do;
%CleanACS(level=Nation,period=2);
%end; /* &nationwide=yes & &oneACSPeriod=no processing */

%if &nationwide. = %quote(yes) %then %do;
/* Clean up the US total demographic data table and keep just what we need*/
data DemogTable_US2;
    length ZoneID $10;
    set DemogTable_US;
    if ZoneID = 'Nationwide';
    format Pct: 6.2;
run;
%end; /* &nationwide=yes processing */

/* Get state-wide pop-weighted means for Census variables */
proc means data=ZoneList2 NOPRINT NWAY MEAN;
    var &keepCens.;
    weight TotalPop;
    output out=StateMeans_Cens MEAN=;
run;
data StateMeans_Cens2;
    length ZoneID $10;
    set StateMeans_Cens;
    ZoneID = 'Statewide';
    drop _TYPE_ _FREQ_;
run;

/* Get statewide pop total (2010) and add to Census variables */
proc means data=ZoneList2 NOPRINT NWAY SUM;
    var TotalPop;
    output out=StatePopTotal SUM=;
run;
data StatePopTotal2;
    length ZoneID $10;
    set StatePopTotal;
    ZoneID = 'Statewide';
    drop _TYPE_ _FREQ_;
run;
proc sort data=StateMeans_Cens2; by ZoneID; run;
proc sort data=StatePopTotal2; by ZoneID; run;
data StateMeans_Cens3;
    merge StatePopTotal2 (in=inPopTot)
        StateMeans_Cens2 (in=inMeans);
    by ZoneID;
    if inPopTot;
    if not inMeans then
    if not inMeans then putlog "*** Missing state-wide means for census variables";
run;

/* Macro to get zone level pop-weighted means for ACS variables, combine with
    Census 2010 variables, and combine zone, state, and national level data  */
/* Set macro parameter for testing:
%let period=per1;
/* */
%macro zoneMeansAndCombine(period=);
/* Add the ZoneID to the tract-level ACS data */
proc sort data=ACSData_&period._Tract2; by GeoID; run;
proc sort data=ZonedTracts2; by TractID; run;
data ACSData_&period._Tract3;
    merge ACSData_&period._Tract2 (in=inACS)
        ZonedTracts2 (in=inZoned);
    by TractID;
    if inZoned;
    if not inACS then putlog "*** Missing &period. ACS data for tract " TractID;
run;
/* Get zone pop-weighted means for ACS variables */
proc means data=ACSData_&period._Tract3 NOPRINT NWAY MEAN;
    class ZoneID;
    weight TotPop;
    output out=ZoneMeans_ACS&period. MEAN=;
run;
data ZoneMeans_ACS&period.b;
    set ZoneMeans_ACS&period.;
    drop _TYPE_ _FREQ_;
    %if &period. = %quote(1) %then %do;
    format &keepACS. 6.2;
    %end;
    %else %do;
    format &keepACS2. 6.2;
    %end;
run;
/* Combine Census and ACS variables by zone */
proc sort data=ZoneList2; by ZoneID; run;
proc sort data=ZoneMeans_ACS&period.b; by ZoneID; run;
data ZoneMeans_CensACS&period.;
    merge ZoneList2 (in=inCens)
        ZoneMeans_ACS&period.b (in=inACS);
    by ZoneID;
    if inCens;
    if not inACS then putlog "*** Missing ACS &period. data for zone " ZoneID;
run;
/* Combine state-wide Census and ACS variables */
proc sort data=StateMeans_Cens3; by ZoneID; run;
proc sort data=ACSData_&period._State2; by ZoneID; run;
data StateMeans_CensACS&period.;
    merge StateMeans_Cens3 (in=inCens)
        ACSData_&period._State2 (in=inACS);
    by ZoneID;
    if inCens;
    if not inACS then putlog "*** Missing ACS &period. data for zone " ZoneID;
run;
%if &nationwide. = %quote(yes) %then %do;
/* Combine nationwide Census and ACS variables */
proc sort data=DemogTable_US2; by ZoneID; run;
proc sort data=ACSData_&period._Nation2; by ZoneID; run;
data NationalMeans_CensACS&period.;
    merge DemogTable_US2 (in=inCens)
        ACSData_&period._Nation2 (in=inACS);
    by ZoneID;
    if inCens;
    if not inACS then putlog "*** Missing ACS &period. data for zone " ZoneID;
run;
%end; /* &nationwide=yes processing */
/* Append state and national level data to the zone-level data */
data DemogVars_CensACS&period.;
    length ZoneID $10;
    set ZoneMeans_CensACS&period.
        StateMeans_CensACS&period.
        %if &nationwide. = %quote(yes) %then %do;
        NationalMeans_CensACS&period.
        %end; /* &nationwide=yes processing */
        ;
run;
%mend zoneMeansAndCombine;

/* Get pop-weighted means for zones and combine with state and national */
%zoneMeansAndCombine(period=per1);
%if &oneACSPeriod. = %quote(no) %then %do;
%zoneMeansAndCombine(period=per2);
%end; /* &oneACSPeriod=no processing */

/* Create 1 year dataset using period 1 ACS data */
data DemogTable_01yr;
    length ZoneID $10 Years $5;
    set DemogVars_CensACSper1;
    Years = "01yr";
    /* QC check PctMinority */
    if PctMinority <= sum(PctHispanic,PctBlackNH,PctAPINH) then
        putlog "Inconsistent race/ethnicity data, ZoneID: " ZoneID;
    keep ZoneID Years TotalPop &keepCens. &keepACS.;
run;

/* Create 5 year dataset using ACS period 1 data (same as 1 year dataset) */
data DemogTable_05yrs;
    length ZoneID $10 Years $5;
    set DemogVars_CensACSper1;
    Years = "05yrs";
    /* QC check PctMinority */
    if PctMinority <= sum(PctHispanic,PctBlackNH,PctAPINH) then
        putlog "Inconsistent race/ethnicity data, ZoneID: " ZoneID;
    keep ZoneID Years TotalPop &keepCens. &keepACS.;
run;

/* Create 10 year dataset averaging ACS period 1 and period 2 data (if available) */
/* Note: if any variables are missing for ACS period 2, will use values from period 1 */
data DemogTable_10yrs;
    length ZoneID $10 Years $5;
    set DemogVars_CensACSper1
        %if &oneACSPeriod. = %quote(no) %then %do;
        DemogVars_CensACSper2
        %end; /* &oneACSPeriod=no processing */
        ;
    Years = "10yrs";
    /* QC check PctMinority */
    if PctMinority <= sum(PctHispanic,PctBlackNH,PctAPINH) then
        putlog "Inconsistent race/ethnicity data, ZoneID: " ZoneID;
    keep ZoneID Years TotalPop &keepCens. &keepACS.;
run;
proc means data=DemogTable_10yrs NOPRINT NWAY MEAN;
    class ZoneID Years;
    output out=DemogTable_10yrs2(drop=_TYPE_ _FREQ_) MEAN=;
run;

/* Append the three pieces for the web tool */
data DemogTable_WebTool;
    set DemogTable_01yr DemogTable_05yrs DemogTable_10yrs2;
    rename ZoneID = Zone;
run;

/* Keep just the data needed for the web tool: filter rows by year */
proc sort data=DemogTable_WebTool; by Years; run;
proc sort data=WebTool_Years; by value; run;
data DemogTable_WebTool2;
    merge DemogTable_WebTool (in=inDemog)
        WebTool_Years (in=inYears keep=value rename=(value=Years));
    by Years;
    if inYears and not inDemog then putlog "*** Unmatched year value in WebTool table:" Years;
    if inDemog and inYears; /* Keep only those rows that are in both datasets */
run;

/* Final sort */
proc sort data=DemogTable_WebTool2; by Zone Years; run;


/* Save SAS dataset */
data ZONEDATA.DemogTable_&stateAbbr.to&ACS_endyr1._WebTool;
    set DemogTable_WebTool2;
run;

/* Export to a CSV file */
proc export data=DemogTable_WebTool2
            OUTFILE= "&pathbase.\DemogTable_&stateAbbr.to&ACS_endyr1._WebTool.csv"
            DBMS=CSV REPLACE;
     PUTNAMES=YES;
run;


/* Summary statistics */
title "52_ZoneWebToolDemog_&stateAbbr.to&ACS_endyr1. - summary statistics";
proc means data=ZONEDATA.DemogTable_&stateAbbr.to&ACS_endyr1._WebTool;
run;

title "52_ZoneWebToolDemog_&stateAbbr.to&ACS_endyr1. - data table";
proc print data=ZONEDATA.DemogTable_&stateAbbr.to&ACS_endyr1._WebTool;
    id Zone;
run;


ods pdf close;

