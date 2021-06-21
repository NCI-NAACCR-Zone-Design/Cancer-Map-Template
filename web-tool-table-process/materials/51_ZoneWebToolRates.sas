/* 51_ZoneWebToolRates - Generate cancer rate table for the web tool from SEER*Stat results */

options sysprintfont=("Courier New" 8) leftmargin=0.75in nocenter compress=no;
ods graphics on;

%let stateAbbr=IA;    /* Set state/registry abbreviation */
%let runNum=IA01; /* Run number used for Step 2 AZTool execution */
%let year1=2015;  /* Latest year */
%let year5=2011_2015;  /* 5-year range */
%let year10=2006_2015;  /* 10-year range */
%let nationwide=yes;  /* Include nationwide rates (yes|no)? */
%let allUSdset=AllUS_Combined_v2_to2015; /* USCS dataset with national rates */

/* Specify data path here for portability: */
%let pathbase=C:\Work\WebToolTables;

libname ZONEDATA "&pathbase.";
ods pdf file="&pathbase.\51_ZoneWebToolRates_&runNum..pdf";

/* Import the final zone list Excel file */
PROC IMPORT OUT=ZoneList
            DATAFILE="&pathbase.\ZoneList_&runNum._final.xlsx"
            DBMS=XLSX REPLACE;
     SHEET="SummaryStats";
run;

/* Import SEER*Stat rate session results - by zone */
PROC IMPORT OUT=SEERin_zones
            DATAFILE="&pathbase.\&runNum.zone_RateCalcs.txt"
            DBMS=DLM REPLACE;
    DELIMITER='09'x; /* Tab */
    GETNAMES=YES;
    DATAROW=2;
    GUESSINGROWS=1000; /* Zone names may be truncated */
run;
/* Import SEER*Stat rate session results - state as a whole */
PROC IMPORT OUT=SEERin_state
            DATAFILE="&pathbase.\&stateAbbr.state_RateCalcs.txt"
            DBMS=DLM REPLACE;
    DELIMITER='09'x; /* Tab */
    GETNAMES=YES;
    DATAROW=2;
    GUESSINGROWS=1000;
run;
%if &nationwide. = %quote(yes) %then %do;
/* Import SEER*Stat rate session results - US as a whole */
/* These results are from a run done by NPCR folks against a CDC-internal database
    and they include 1, 5, and 10 year results */
PROC IMPORT OUT=SEERin_allUS
            DATAFILE="&pathbase.\National_Cancer_Rates\&allUSdset..txt"
            DBMS=DLM REPLACE;
    DELIMITER='09'x; /* Tab */
    GETNAMES=YES;
    DATAROW=2;
    GUESSINGROWS=1000;
run;
%end; /* &nationwide=yes processing */

/* Import SEER*Stat cancer site table */
PROC IMPORT OUT=SeerStatSites
            DATAFILE="&pathbase.\CancerSiteTable_&stateAbbr..xlsx"
            DBMS=XLSX REPLACE;
     SHEET="SEER_Stat";
run;

/* Import WebTool tables for site, years, and races */
PROC IMPORT OUT=WebTool_CancerSites
            DATAFILE="&pathbase.\CancerSiteTable_&stateAbbr..xlsx"
            DBMS=XLSX REPLACE;
     SHEET="Webtool CANCERSITE";
run;
PROC IMPORT OUT=WebTool_Years
            DATAFILE="&pathbase.\CancerSiteTable_&stateAbbr..xlsx"
            DBMS=XLSX REPLACE;
     SHEET="Webtool TIME";
run;
PROC IMPORT OUT=WebTool_RaceEthGroups
            DATAFILE="&pathbase.\CancerSiteTable_&stateAbbr..xlsx"
            DBMS=XLSX REPLACE;
     SHEET="Webtool RACE";
run;

/* Clear formats, informats and labels for imported datasets */
proc datasets lib=work nolist;
    MODIFY ZoneList; FORMAT _char_; INFORMAT _char_; ATTRIB _all_ label=''; run;
    MODIFY SEERin_zones; FORMAT _all_; INFORMAT _all_; ATTRIB _all_ label=''; run;
    MODIFY SEERin_state; FORMAT _all_; INFORMAT _all_; ATTRIB _all_ label=''; run;
    %if &nationwide. = %quote(yes) %then %do;
        MODIFY SEERin_allUS; FORMAT _all_; INFORMAT _all_; ATTRIB _all_ label=''; run;
    %end; /* &nationwide=yes processing */
    MODIFY SeerStatSites; FORMAT _all_; INFORMAT _all_; ATTRIB _all_ label=''; run;
    MODIFY WebTool_CancerSites; FORMAT _all_; INFORMAT _all_; ATTRIB _all_ label=''; run;
    MODIFY WebTool_Years; FORMAT _all_; INFORMAT _all_; ATTRIB _all_ label=''; run;
    MODIFY WebTool_RaceEthGroups; FORMAT _all_; INFORMAT _all_; ATTRIB _all_ label=''; run;
quit;


/* Extract ZoneID and possibly truncated ZoneName */
data SEERin_zones2;
    length ZoneID $10 ZoneNameTrunc $200;
    set SEERin_zones;
    ZoneID = scan(Zones_&runNum._final,1,':');
    ZoneNameTrunc = scan(Zones_&runNum._final,2,':');
run;

/* Add full zone name */
proc sort data=SEERin_zones2; by ZoneID; run;
proc sort data=ZoneList; by ZoneIDOrig; run;
data SEERin_zones3;
    length ZoneID $10 ZoneName $200;
    merge SEERin_zones2 (in=inData)
        ZoneList (in=inNames keep=ZoneName ZoneIDOrig
            rename=(ZoneIDOrig=ZoneID));
    by ZoneID;
    if inData;
    if not inNames then putlog "*** Missing zone name for ZoneID: " ZoneID;
    drop ZoneNameTrunc Zones_&runNum._final;
run;

/* Add a ZoneID variable to the state dataset so we can combine with zones */
data SEERin_state2;
    length ZoneID $10;
    set SEERin_state;
    ZoneID = "Statewide";
run;

%if &nationwide. = %quote(yes) %then %do;
/* Add a ZoneID variable to the US dataset so we can combine with zones */
data SEERin_allUS2;
    length ZoneID $10;
    set SEERin_allUS;
    ZoneID = "Nationwide";
run;
%end; /* &nationwide=yes processing */

/* Combine zone and state datasets */
data RateTable;
    set SEERin_zones3 SEERin_state2;
run;

/* Clean up the data and rename columns */
data RateTable2;
    length ZoneID $10 SexNew $10 Site $40 Years $5 RaceEth $12;
    set RateTable;
    if Sex = 'Male and female' then SexNew = 'Both';
    else SexNew = Sex;
    Site = USCS_Sites;
    select (LatestYears_1_5_10);
        when ("1yr_&year1.")    Years='01yr';
        when ("5yrs_&year5.")   Years='05yrs';
        when ("10yrs_&year10.") Years='10yrs';
        otherwise               Years='???';
        end;
    /* Set RaceEth variable and delete unknowns */
    select (Race_and_origin_recode_with_All);
        when ('AllRaceEth')                                 RaceEth='.AllRaceEth';
        when ('Non-Hispanic White')                         RaceEth='White_NH';
        when ('Non-Hispanic Black')                         RaceEth='Black_NH';
        when ('Non-Hispanic Asian or Pacific Islander')     RaceEth='API_NH';
        when ('Non-Hispanic American Indian/Alaska Native') RaceEth='AIAN_NH';
        when ('Non-Hispanic Unknown Race')                  RaceEth='Unknown_NH';
        when ('Hispanic (All Races)')                       RaceEth='Hispanic';
        otherwise                                           RaceEth='???';
        end;
    if RaceEth = 'Unknown_NH' then delete;
    drop Sex USCS_Sites LatestYears_1_5_10 Race_and_origin_recode_with_All Standard_Error;
    rename
        SexNew = Sex
        Age_Adjusted_Rate = AAIR
        Lower_Confidence_Interval = LCI
        Upper_Confidence_Interval = UCI
        Count = Cases
        Population = PopTot;
run;

%if &nationwide. = %quote(yes) %then %do;
/* Clean up the US data and rename columns to match */
/* (Because these data come from the CDC USCS SEER*Stat database,
    race and ethnicity are in separate variables.) */
data RateTable_US;
    length ZoneID $10 SexNew $10 Site $40 Years $5 RaceEth $12;
    set SEERin_allUS2;
    if index(LatestYears_1_5_10,"&year1.") > 0; /* Keep just the years ending in the &year1. value */
    if Sex = 'Male and female' then SexNew = 'Both';
    else SexNew = Sex;
    Site = USCS_Sites;
    select (LatestYears_1_5_10);
        when ("1yr_&year1.")    Years='01yr';
        when ("5yrs_&year5.")   Years='05yrs';
        when ("10yrs_&year10.") Years='10yrs';
        otherwise               Years='???';
        end;
    /* Set RaceEth variable and delete unneeded combinations */
    if      (Race_recode_for_uscs = 'All races') and
            (Origin_recode_with_AllOrigins = 'AllOrigins')
                then RaceEth='.AllRaceEth';
    else if (Race_recode_for_uscs = 'White') and
            (Origin_recode_with_AllOrigins = 'Non-Spanish-Hispanic-Latino')
                then RaceEth='White_NH';
    else if (Race_recode_for_uscs = 'Black') and
            (Origin_recode_with_AllOrigins = 'Non-Spanish-Hispanic-Latino')
                then RaceEth='Black_NH';
    else if (Race_recode_for_uscs = 'Asian or Pacific Islander') and
            (Origin_recode_with_AllOrigins = 'Non-Spanish-Hispanic-Latino')
                then RaceEth='API_NH';
    else if (Race_recode_for_uscs = 'American Indian/Alaska Native') and
            (Origin_recode_with_AllOrigins = 'Non-Spanish-Hispanic-Latino')
                then RaceEth='AIAN_NH';
    else if (Race_recode_for_uscs = 'All races') and
            (Origin_recode_with_AllOrigins = 'Spanish-Hispanic-Latino')
                then RaceEth='Hispanic';
    else delete; /* Delete unneeded combinations */
    drop Sex USCS_Sites LatestYears_1_5_10 Race_recode_for_uscs Origin_recode_with_AllOrigins Standard_Error;
    rename
        SexNew = Sex
        Age_Adjusted_Rate = AAIR
        Lower_Confidence_Interval = LCI
        Upper_Confidence_Interval = UCI
        Count = Cases
        Population = PopTot;
run;
%end; /* &nationwide=yes processing */

/* Add the US data to the main rate table */
data RateTable3;
    set RateTable2
        %if &nationwide. = %quote(yes) %then %do;
        RateTable_US
        %end; /* &nationwide=yes processing */
        ;
run;

/* Modify sex-specific site names and remove opposite sex observations */
proc sort data=RateTable3; /* Sort by Site and Sex last so we can verify the changes */
    by ZoneID Years RaceEth Site Sex;
run;
data RateTable4;
    set RateTable3;
    SexSpecSite = 0; /* Sex-specific site flag */
    if Site = 'Breast' then do;
        SexSpecSite = 1;
        if Sex = 'Female' then Site = 'Breast (female)';
        else delete;
        end;
    if Site = 'Cervix Uteri' then do;
        SexSpecSite = 1;
        if Sex = 'Female' then Site = 'Cervix Uteri (female)';
        else delete;
        end;
    if Site = 'Corpus and Uterus, NOS' then do;
        SexSpecSite = 1;
        if Sex = 'Female' then Site = 'Corpus and Uterus, NOS (female)';
        else delete;
        end;
    if Site = 'Ovary' then do;
        SexSpecSite = 1;
        if Sex = 'Female' then Site = 'Ovary (female)';
        else delete;
        end;
    if Site = 'Prostate' then do;
        SexSpecSite = 1;
        if Sex = 'Male' then Site = 'Prostate (male)';
        else delete;
        end;
    if Site = 'Testis' then do;
        SexSpecSite = 1;
        if Sex = 'Male' then Site = 'Testis (male)';
        else delete;
        end;
run;

/* Create a cancer site sort sequence field based on state rates */
data SiteSortSeq;
    set RateTable4;
    if ZoneID = "Statewide";
    if Years = '10yrs';
    if RaceEth = '.AllRaceEth';
    if (Sex = 'Female') and (index(Site,'female')=0) then delete;
    if (Sex = 'Male') and (index(Site,'male')=0) then delete;
    /* Adjust rate for sex-specific cancer sites */
    if (Sex = 'Female') or (Sex = 'Male') then AAIR = AAIR / 2;
    keep Site AAIR;
run;
proc sort data=SiteSortSeq; by descending AAIR; run;
data SiteSortSeq2;
    set SiteSortSeq;
    SiteSort = _N_;
run;

/* Add the site sort sequence field to the rate table */
proc sort data=RateTable4; by Site; run;
proc sort data=SiteSortSeq2; by Site; run;
data RateTable5;
    merge RateTable4 (in=inRates)
        SiteSortSeq2 (in=inSort drop=AAIR);
    by Site;
    if inRates;
    if not inSort then putlog "*** Missing sort sequence number: " ZoneID Years RaceEth Site Sex;
    rename Site = Site_full;
run;

/* Add short site name */
data SeerStatSites2;
    length Site_short $10 Site_SEERStat_var $40 Site_sex $6 Site_full $40;
    set SeerStatSites;
    if Site_sex ^= '' then Site_full = catt(Site_SEERStat_var, ' (', Site_sex, ')');
    else Site_full = Site_SEERStat_var;
run;
proc sort data=RateTable5; by Site_full; run;
proc sort data=SeerStatSites2; by Site_full; run;
data RateTable6;
    length ZoneID $10 Sex $10 Site_short $10 Years $5 RaceEth $12;
    merge RateTable5 (in=inRates)
        SeerStatSites2 (in=inSites);
    by Site_full;
    if inRates;
    if not inSites then putlog "*** Unexpected missing site name: " Site_full;
    drop Site_SEERStat_var Site_sex Site_full;
run;

/* Suppress counts and rates if 15 or fewer cases */
data RateTable_wSuppr;
    set RateTable6;
    if Cases < 16 then do;
        Cases = .;
        AAIR = .;
        LCI = .;
        UCI = .;
        end;
    if Cases = . then LT16cases = 1; /* For summary suppression statistics */
    else LT16cases = 0;
    /* Add a ByGroup variable for summary suppression statistics */
    length ByGroup $18;
    if (Sex = 'Both') and (RaceEth = '.AllRaceEth') then ByGroup = '1-BySite';
    if (Sex ^= 'Both') and (RaceEth = '.AllRaceEth') then do;
        if SexSpecSite = 1 then ByGroup = '1-BySite';
        else                    ByGroup = '2-BySiteSex';
        end;
    if (Sex = 'Both') and (RaceEth ^= '.AllRaceEth') then ByGroup = '3-BySiteRaceEth';
    if (Sex ^= 'Both') and (RaceEth ^= '.AllRaceEth') then do;
        if SexSpecSite = 1 then ByGroup = '3-BySiteRaceEth';
        else                    ByGroup = '4-BySiteSexRaceEth';
        end;
    drop SexSpecSite;
run;

/* Create separate rate variables for each race/ethnicity */
proc sort data=RateTable_wSuppr;
    by ZoneID Site_short Sex Years RaceEth;
run;
data RateTable_WebTool;
    length ZoneID $10 Sex $10 Site_short $10 Years $5;
    retain
        All_PopTot All_Cases All_AAIR All_LCI All_UCI .
        W_PopTot W_Cases W_AAIR W_LCI W_UCI .
        B_PopTot B_Cases B_AAIR B_LCI B_UCI .
        H_PopTot H_Cases H_AAIR H_LCI H_UCI .
        API_PopTot API_Cases API_AAIR API_LCI API_UCI .
        AIAN_PopTot AIAN_Cases AIAN_AAIR AIAN_LCI AIAN_UCI .
        ;
    set RateTable_wSuppr;
    by ZoneID Site_short Sex Years;
    select (RaceEth);
        when ('.AllRaceEth') do;
            All_PopTot = PopTot; All_Cases = Cases; All_AAIR = AAIR; All_LCI = LCI; All_UCI = UCI;
            end;
        when ('White_NH') do;
            W_PopTot = PopTot; W_Cases = Cases; W_AAIR = AAIR; W_LCI = LCI; W_UCI = UCI;
            end;
        when ('Black_NH') do;
            B_PopTot = PopTot; B_Cases = Cases; B_AAIR = AAIR; B_LCI = LCI; B_UCI = UCI;
            end;
        when ('API_NH') do;
            API_PopTot = PopTot; API_Cases = Cases; API_AAIR = AAIR; API_LCI = LCI; API_UCI = UCI;
            end;
        when ('AIAN_NH') do;
            AIAN_PopTot = PopTot; AIAN_Cases = Cases; AIAN_AAIR = AAIR; AIAN_LCI = LCI; AIAN_UCI = UCI;
            end;
        when ('Hispanic') do;
            H_PopTot = PopTot; H_Cases = Cases; H_AAIR = AAIR; H_LCI = LCI; H_UCI = UCI;
            end;
        otherwise putlog "*** Unexpected race ethnicity value: " RaceEth ZoneID Site_short Sex Years;
        end;
    if last.Years then do;
        output;
        All_PopTot=.; All_Cases=.; All_AAIR=.; All_LCI=.; All_UCI=.;
        W_PopTot=.; W_Cases=.; W_AAIR=.; W_LCI=.; W_UCI=.;
        B_PopTot=.; B_Cases=.; B_AAIR=.; B_LCI=.; B_UCI=.;
        H_PopTot=.; H_Cases=.; H_AAIR=.; H_LCI=.; H_UCI=.;
        API_PopTot=.; API_Cases=.; API_AAIR=.; API_LCI=.; API_UCI=.;
        AIAN_PopTot=.; AIAN_Cases=.; AIAN_AAIR=.; AIAN_LCI=.; AIAN_UCI=.;
        end;
    format
        All_AAIR All_LCI All_UCI
        W_AAIR W_LCI W_UCI
        B_AAIR B_LCI B_UCI
        H_AAIR H_LCI H_UCI
        API_AAIR API_LCI API_UCI
        AIAN_AAIR AIAN_LCI AIAN_UCI
        8.1;
    drop RaceEth AAIR LCI UCI Cases PopTot
        LT16cases ByGroup;
    rename
        ZoneID = Zone
        Site_short = Cancer
        All_PopTot = PopTot
        All_Cases = Cases
        All_AAIR = AAIR
        All_LCI = LCI
        All_UCI = UCI;
run;

/* Keep just the data needed for the web tool: filter rows by cancer site */
proc sort data=RateTable_WebTool; by Cancer; run;
proc sort data=WebTool_CancerSites; by value; run;
data RateTable_WebTool2;
    merge RateTable_WebTool (in=inRates)
        WebTool_CancerSites (in=inSites keep=value rename=(value=Cancer));
    by Cancer;
    if inSites and not inRates then putlog "*** Unmatched short site name in WebTool table:" Cancer;
    if inRates and inSites; /* Keep only those rows that are in both datasets */
run;

/* Keep just the data needed for the web tool: filter rows by year */
proc sort data=RateTable_WebTool2; by Years; run;
proc sort data=WebTool_Years; by value; run;
data RateTable_WebTool3;
    merge RateTable_WebTool2 (in=inRates)
        WebTool_Years (in=inYears keep=value rename=(value=Years));
    by Years;
    if inYears and not inRates then putlog "*** Unmatched year value in WebTool table:" Years;
    if inRates and inYears; /* Keep only those rows that are in both datasets */
run;

/* Keep just the data needed for the web tool: drop unneeded race/eth variables */
data _null_; /* Create a macro variable for each race/eth group */
    set WebTool_RaceEthGroups;
    if _N_ = 1 then do;
        call symputx("KeepWhite", "No");
        call symputx("KeepBlack", "No");
        call symputx("KeepHisp", "No");
        call symputx("KeepAPI", "No");
        call symputx("KeepAIAN", "No");
        end;
    if value = "W" then call symputx("KeepWhite", "Yes");
    if value = "B" then call symputx("KeepBlack", "Yes");
    if value = "H" then call symputx("KeepHisp", "Yes");
    if value = "API" then call symputx("KeepAPI", "Yes");
    if value = "AIAN" then call symputx("KeepAIAN", "Yes");
run;
data _null_; /* Verify macro variable values */
    putlog "KeepWhite: &KeepWhite.";
    putlog "KeepBlack: &KeepBlack.";
    putlog "KeepHisp: &KeepHisp.";
    putlog "KeepAPI: &KeepAPI.";
    putlog "KeepAIAN: &KeepAIAN.";
run;
data RateTable_WebTool4; /* Drop unneeded variables */
    set RateTable_WebTool3;
    %if &KeepWhite. = %quote(No) %then %do;
        drop W_:; /* Drop the five White cancer rate variables */
        %end;
    %if &KeepBlack. = %quote(No) %then %do;
        drop B_:; /* Drop the five Black cancer rate variables */
        %end;
    %if &KeepHisp. = %quote(No) %then %do;
        drop H_:; /* Drop the five Hispanic cancer rate variables */
        %end;
    %if &KeepAPI. = %quote(No) %then %do;
        drop API_:; /* Drop the five API cancer rate variables */
        %end;
    %if &KeepAIAN. = %quote(No) %then %do;
        drop AIAN_:; /* Drop the five AIAN cancer rate variables */
        %end;
run;

/* Final sorts */
%macro FinalSort(ds=,zoneIdVar=,lastBy=);
proc sort data=&ds.;
    by Years SiteSort Sex &zoneIdVar. &lastBy.;
run;
%mend FinalSort;
%FinalSort(ds=RateTable_wSuppr,zoneIdVar=ZoneID,lastBy=RaceEth);
%FinalSort(ds=RateTable_WebTool4,zoneIdVar=Zone,lastBy=);


/* Generate summary tables of percent suppressed */

/* Create a dataset with just the zone rates (no state or US rates) */
data RateTable_wSupprZonesOnly;
    set RateTable_wSuppr;
    if ZoneID in ("Statewide", "Nationwide") then delete;
run;

/* Calculate percent suppressed and clean up */
%MACRO CalcPct(classVars=, suffix=);
proc summary data=RateTable_wSupprZonesOnly noprint nway;
    var LT16cases;
    class &classVars. ByGroup Years;
    output out=Summ&suffix. SUM= / AUTONAME AUTOLABEL;
run;
data Summ&suffix.2;
    set Summ&suffix.;
    if LT16cases_Sum = . then LT16cases_Sum = 0;
    PctZoneSuppr = 100 * LT16cases_Sum / _FREQ_;
    format PctZoneSuppr 7.2;
    drop _TYPE_;
    rename _FREQ_ = NumCells
        LT16cases_Sum = SupprCells;
run;
data Summ&suffix.3;
    retain PctSuppr_1yr PctSuppr_5yrs PctSuppr_10yrs .;
    set Summ&suffix.2;
    by &classVars. ByGroup;
    if Years = '01yr' then PctSuppr_1yr = PctZoneSuppr;
    if Years = '05yrs' then PctSuppr_5yrs = PctZoneSuppr;
    if Years = '10yrs' then PctSuppr_10yrs = PctZoneSuppr;
    if last.ByGroup then do;
        output;
        PctSuppr_1yr = .;
        PctSuppr_5yrs = .;
        PctSuppr_10yrs = .;
        end;
    format PctSuppr_1yr PctSuppr_5yrs PctSuppr_10yrs 7.2;
    drop Years SupprCells PctZoneSuppr;
run;
proc sort data=Summ&suffix.3; by ByGroup; run;
%MEND CalcPct;
%CalcPct(classVars=, suffix=All);
%CalcPct(classVars=RaceEth, suffix=Race);
%CalcPct(classVars=SiteSort Site_short, suffix=Site);
%CalcPct(classVars=RaceEth SiteSort Site_short, suffix=RaceSite);

/* Delete the '.AllRaceEth' rows from the by-race-and-site dataset */
data SummRaceSite4;
    set SummRaceSite3;
    if RaceEth = '.AllRaceEth' then delete;
run;
proc sort data=SummRaceSite4; by ByGroup descending RaceEth; run;

/* Set up to export selected suppression summary tables to Excel */
data ExcelSumm_BySite
     ExcelSumm_BySiteSex;
    length Site_short $10 NumCells 8;
    set SummSite3;
    if ByGroup = '1-BySite'    then output ExcelSumm_BySite;
    if ByGroup = '2-BySiteSex' then output ExcelSumm_BySiteSex;
    drop SiteSort ByGroup;
run;
data ExcelSumm_BySiteWhiteNH
     ExcelSumm_BySiteHisp
     ExcelSumm_BySiteBlackNH
     ExcelSumm_BySiteAPINH
     ExcelSumm_BySiteAIANNH;
    length RaceEth $12 Site_short $10 NumCells 8;
    set SummRaceSite4;
    if ByGroup = '3-BySiteRaceEth' then do;
        if RaceEth = 'White_NH' then output ExcelSumm_BySiteWhiteNH;
        if RaceEth = 'Hispanic' then output ExcelSumm_BySiteHisp;
        if RaceEth = 'Black_NH' then output ExcelSumm_BySiteBlackNH;
        if RaceEth = 'API_NH' then output ExcelSumm_BySiteAPINH;
        if RaceEth = 'AIAN_NH' then output ExcelSumm_BySiteAIANNH;
        end;
    drop SiteSort ByGroup;
run;


/* Save the main SAS datasets */
data ZONEDATA.RateTable_&stateAbbr._wSuppr;  /* Rates with suppression */
    set RateTable_wSuppr;
run;
data ZONEDATA.RateTable_&stateAbbr._WebTool;  /* Rates for WebTool */
    set RateTable_WebTool4;
run;

/* Export web tool dataset to a CSV files */
proc export data=RateTable_WebTool (drop=SiteSort)
            OUTFILE= "&pathbase.\RateTable_&stateAbbr._WebTool.csv"
            DBMS=csv REPLACE;
     PUTNAMES=YES;
run;

/* Export selected suppression summary tables to Excel */
%MACRO ExportExcel(dset=);
proc export data=ExcelSumm_&dset.
            OUTFILE= "&pathbase.\RateTable_&stateAbbr._SupprSumm.xlsx"
            DBMS=XLSX REPLACE;
     SHEET="&dset.";
run;
%MEND ExportExcel;
%ExportExcel(dset=BySite);
%ExportExcel(dset=BySiteSex);
%ExportExcel(dset=BySiteWhiteNH);
%ExportExcel(dset=BySiteBlackNH);
%ExportExcel(dset=BySiteHisp);
%ExportExcel(dset=BySiteAPINH);
%ExportExcel(dset=BySiteAIANNH);
===
/* Export selected suppression summary tables to Excel */
/* Delete the target Excel file if it exists */
data _null_;
    fname = 'todelete';
    rc = filename(fname, "&pathbase.\RateTable_&stateAbbr._SupprSumm.xlsx");
    if rc = 0 and fexist(fname) then do;
        rc = fdelete(fname);
        if rc > 0 then putlog "*** Failed to delete previous Excel file, rc=" rc;
        end;
    rc = filename(fname);
run;
/* Assign a libname to the target Excel file */
libname OutExcel "&pathbase.\RateTable_&stateAbbr._SupprSumm.xlsx";
/* Macro to create a worksheet */
%MACRO ExportExcel(dset=);
data OutExcel.&dset.; /* Worksheet = &dset. */
    set ExcelSumm_&dset.;
run;
%MEND ExportExcel;
/* Create worksheets */
%ExportExcel(dset=BySite);
%ExportExcel(dset=BySiteSex);
%ExportExcel(dset=BySiteWhiteNH);
%ExportExcel(dset=BySiteBlackNH);
%ExportExcel(dset=BySiteHisp);
%ExportExcel(dset=BySiteAPINH);
%ExportExcel(dset=BySiteAIANNH);
/* Close the Excel file */
libname OutExcel;


/* Summary statistics */
title "51_ZoneWebToolRates_&runNum. - summary statistics for web tool rate table";
proc freq data=ZONEDATA.RateTable_&stateAbbr._WebTool;
    table Cancer / list missing;
    table Sex / list missing;
    table Years / list missing;
run;
ods pdf STARTPAGE=NO;
proc means data=ZONEDATA.RateTable_&stateAbbr._WebTool;
run;
ods pdf STARTPAGE=YES;

title "51_ZoneWebToolRates_&runNum. - summary statistics for original rates with suppression";
proc freq data=ZONEDATA.RateTable_&stateAbbr._wSuppr;
    table Site_short / list missing;
    table Sex / list missing;
    table Years / list missing;
    table RaceEth / list missing;
    table LT16cases / list missing;
    table ByGroup*LT16cases / list missing;
    table Years*ByGroup*LT16cases / list missing;
run;
ods pdf STARTPAGE=NO;
proc means data=ZONEDATA.RateTable_&stateAbbr._wSuppr;
    var AAIR LCI UCI Cases PopTot;
run;
ods pdf STARTPAGE=YES;

/* Print the zone suppression summary tables */
title "51_ZoneWebToolRates_&runNum. - zone suppression summary tables";
proc print data=SummAll3 split='_';
    ID ByGroup;
    var NumCells PctSuppr_1yr PctSuppr_5yrs PctSuppr_10yrs;
run;
proc print data=SummRace3 split='_';
    by ByGroup;
    ID RaceEth;
    var NumCells PctSuppr_1yr PctSuppr_5yrs PctSuppr_10yrs;
run;
proc print data=SummSite3 split='_';
    by ByGroup;
    ID Site_short;
    var NumCells PctSuppr_1yr PctSuppr_5yrs PctSuppr_10yrs;
run;
proc print data=SummRaceSite4 split='_';
    by ByGroup descending RaceEth;
    ID RaceEth Site_short;
    var NumCells PctSuppr_1yr PctSuppr_5yrs PctSuppr_10yrs;
run;


ods pdf close;

