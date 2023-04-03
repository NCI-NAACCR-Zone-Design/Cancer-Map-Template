/* ZD_DeleteFile - macro to delete a file if it exists */
%MACRO ZD_DeleteFile(delfile=);
data _null_;
    fname = 'todelete';
    rc = filename(fname, "&delfile.");
    if rc = 0 and fexist(fname) then do;
        rc = fdelete(fname);
        if rc > 0 then putlog "*** Failed to delete file: &delfile., rc=" rc;
        end;
    rc = filename(fname);
run;
%MEND ZD_DeleteFile;
