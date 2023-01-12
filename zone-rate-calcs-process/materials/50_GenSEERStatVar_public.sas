/* 50_GenSEERStatVar - generate a SEER*Stat user variable for the final zone configuration */

/* 5/25/2022 - added a Zone Design AUTOCALL macro library for commonly used macros */
/* 5/25/2022 - modified Excel file imports to use new ZD_ImportExcel macro */
/* 6/24/2022 - added the state FIPS code to the ZoneID so that it matches the combined SEER*Stat database */
/* 11/17/2022 - modified to use ZoneIDFull as the zone ID rather than "Zones_&runNum._final" */

options sysprintfont=("Courier New" 8) leftmargin=0.75in nocenter compress=no;
ods graphics on;

%let stateName=Iowa;  /* Set state/registry name */
%let stateAbbr=IA;    /* Set state/registry abbreviation */
%let runNum=IA01;     /* Run number used for the final zone design */

/* Specify data paths here for portability: */
%let pathbase=C:\Work\ZoneRateCalcs;

libname ZONEDATA "&pathbase.";

/* Set up the Zone Design AUTOCALL macro library */
filename ZDAUTOS "&pathbase.\ZDMacros";
options mautosource sasautos=(SASAUTOS ZDAUTOS);

/* Import the final zoned tracts Excel file */
%ZD_ImportExcel(sourcefile=&pathbase.\ZonedTracts_&runNum._final.xlsx,
    sheetname=Tracts, /*11/23/22 Updated sheetname*/
    targetds=ZonedTracts);

/* Import the final zone list Excel file */
%ZD_ImportExcel(sourcefile=&pathbase.\ZoneList_&runNum._final.xlsx,
    sheetname=FullStats, /*11/23/22 Updated sheetname*/
    targetds=ZoneList);

/* Read in the SEER*Stat tract file for this state */
data SeerStat_TractFile;
    infile "&pathbase.\Tracts_&stateAbbr..fmx"
        length=inLength truncover;
    input @1 Content $100.;
    InLen = inLength;
run;

/* Parse SEER*Stat tract file */
data SeerStat_Tracts;
    length SeerTractNum 8 TractID $11 TractFieldName $100;
    retain TractListFlag 0 TractFieldFlag 0 TractFieldName "";
    set SeerStat_TractFile;
    if InLen = 0 then do;
        TractListFlag = 0;
        TractFieldFlag = 0;
        end;
    if TractListFlag = 1 then do;
        SeerTractNum = input(scan(Content,1,'='),BEST.);
        TractID = scan(Content,2,'=');
        output;
        end;
    if TractFieldFlag = 1 then do; /* Capture the name of the tract field in the SEER*Stat database */
        TractFieldName = scan(Content,2,'=');
        call symputx("SeerStatTractField", TractFieldName);
        end;
    if Content = "[Format=Tracts_&stateAbbr.]" then TractListFlag = 1;
    if Content = "[Field Links]" then TractFieldFlag = 1; /* Next record has tract field name */
    keep SeerTractNum TractID;
run;

/* Add SEER*Stat tract number to Zoned Tracts */
proc sort data=ZonedTracts; by TractID; run;
proc sort data=SeerStat_Tracts; by TractID; run;
data ZonedTracts2;
    merge ZonedTracts (in=inTracts)
        SeerStat_Tracts (in=inSeerStat);
    by TractID;
    if inTracts;
    /* Report any missing tracts that have non-zero pops - should not be any */
    if (not inSeerStat) and (Pop2010 > 0) then putlog "*** Missing SEER*Stat tract, TractID: " TractID
        " ZoneID:" ZoneIDFull
        " Pop2010:" Pop2010;
    keep TractID ZoneIDFull SeerTractNum Pop2010;
run;

/* Create a SEER*Stat fmx file */
%let maxTrkListLen = 250;
%let maxTempOutLen = 300;

/* Generate list of SEER*Stat tract numbers for each zone and add the state FIPS code to the ZoneID */
proc sort data=ZonedTracts2; by ZoneIDFull; run;
data ZoneTractLists;
    length ZoneIDFull $10 TractCnt 8 TractList $&maxTrkListLen.;
    retain TractCnt 0 TractList '' MissCnt 0;
    set ZonedTracts2;
    by ZoneIDFull;
    if SeerTractNum ^= . then do;
        TractList = catx(',',TractList,put(SeerTractNum,BEST.));
        if length(TractList) > &maxTrkListLen. - 5 then
            putlog "*** TractList length close to &maxTrkListLen. for ZoneIDFull: "
                ZoneIDFull;
        TractCnt = TractCnt + 1;
        end;
    else do; /* SeerTractNum = . */
        if Pop2010 = 0 then MissCnt = MissCnt + 1;
        end;
    if last.ZoneIDFull then do;
        output;
        TractCnt = 0;
        TractList = '';
        MissCnt = 0;
        end;
    keep ZoneIDFull TractCnt TractList MissCnt;
run;

/* Add ZoneTractCnt and verify */ /*11/23/22 Updated reference of ZoneTctCnt to ZoneTractCnt*/
proc sort data=ZoneTractLists; by ZoneIDFull; run;
proc sort data=ZoneList;  by ZoneIDFull; run;
data ZoneTractLists2;
    merge ZoneTractLists (in=inTrList)
        ZoneList (in=inZoList keep=ZoneIDFull ZoneTractCnt);
    by ZoneIDFull;
    if inTrList;
    if not inZoList then putlog "*** Missing ZoneList entry for ZoneID: " ZoneIDFull;
    if ZoneTractCnt ^= TractCnt + MissCnt then
        putlog "*** Unexpected difference in zone tract counts, ZoneID:" ZoneIDFull
            " Orig count:" ZoneTractCnt " Generated count:" TractCnt;
run;
proc sort data=ZoneTractLists2; by ZoneIDFull; run;

/* Create a SEER*Stat fmx file */
data _null_;
    length tempout $&maxTempOutLen.;
    set ZoneTractLists2 end=eof;
    file "&pathbase.\ZoneID_&stateAbbr..fmx";
    if _N_ = 1 then do; /* Intro text*/
        put "[Field Links]";
        put "ZoneID_&stateAbbr.=&SeerStatTractField.";
        put;
        put "[FormatDataType]";
        put "ZoneID_&stateAbbr.=I";
        put;
        put "[FormatEditableByUser]";
        put "ZoneID_&stateAbbr.=true";
        put;
        put "[FormatCaseSensitivity]";
        put "ZoneID_&stateAbbr.=true";
        put;
        put "[Format=ZoneID_&stateAbbr.]";
        end;
    tempout = catx('=',TractList,ZoneIDFull);
    put tempout;
run;


/* End of program */
