/* ZD_ExportExcel - macro to export a SAS dataset to an Excel worksheet */
%macro ZD_ExportExcel(sourceds=,targetfile=,sheetname=);
proc export DATA=&sourceds.
            OUTFILE="&targetfile."
            DBMS=XLSX REPLACE;
    SHEET="&sheetname.";
run;
/* Delete the ".bak" file if it has been created */
%ZD_DeleteFile(delfile=&targetfile..bak);
%mend ZD_ExportExcel;
