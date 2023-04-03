/* ZD_ImportExcel - macro to import an Excel worksheet to a SAS dataset */
%macro ZD_ImportExcel(sourcefile=,sheetname=,targetds=);
proc import OUT=&targetds.
            DATAFILE="&sourcefile."
            DBMS=XLSX REPLACE;
    SHEET="&sheetname.";
run;
/* Clear formats informats and labels */
proc datasets lib=WORK nolist;
    MODIFY &targetds.; FORMAT _char_; INFORMAT _char_; ATTRIB _all_ label=''; run;
quit;
%mend ZD_ImportExcel;

