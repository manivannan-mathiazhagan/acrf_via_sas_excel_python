/*****************************************************************************************************/
/* Macro Name     : add_bookmarks                                                                    */
/*                                                                                                   */
/* Purpose        : Automates the creation of PDF bookmarks and an optional Table of Contents (TOC)  */
/*                  using PDFtk and Python.                                                          */
/*                                                                                                   */
/* Author         : Manivannan Mathialagan                                                           */
/* Created On     : 30-Jan-2023                                                                      */
/*                                                                                                   */
/* Parameters                                                                                        */
/*   - FILE_PATH   (Optional) : Directory path for input PDF and output files. Defaults to the       */
/*                              directory where the macro is stored.                                 */
/*                                                                                                   */
/*   - INPDF       (Required) : Name of the input PDF file (without extension).                      */
/*                                                                                                   */
/*   - OUTPDF      (Optional) : Name of the output PDF file. Default is 'aCRF'.                      */
/*                                                                                                   */
/*   - Create_TOC  (Optional) : Flag to indicate whether TOC should be created (Y/N). Default is 'N'.*/
/*                                                                                                   */
/*   - TOC_SIZE    (Optional) : Font size for TOC text. Default is 12.                               */
/*                                                                                                   */
/* Example Usage                                                                                     */
/*   %add_bookmarks(INPDF=blankCRF_Positioned, Create_TOC=Y);                                        */
/*                                                                                                   */
/* Notes                                                                                             */
/*   - This macro requires both Python and PDFtk to be installed and accessible via the system PATH. */
/*   - Ensure proper file paths are used, especially if they contain spaces or special characters.   */
/*   - TOC page numbers may require manual adjustment based on the structure of the final document.  */
/*****************************************************************************************************/

%macro add_bookmarks(FILE_PATH, INPDF=, OUTPDF=aCRF, Create_TOC=N, TOC_SIZE=12);

    /* Check if FILE_PATH is provided, otherwise use the path where the program is stored */
    %if "&FILE_PATH." eq "" %then 
    	%do;
	        %let fullpath = %sysget(SAS_EXECFILEPATH);

	        /* Remove the file name, keep only the directory */
	        %let FILE_PATH = %substr(&fullpath, 1, %length(&fullpath) - %length(%scan(&fullpath, -1, '\')) - 1);
	        %put Full path: &FILE_PATH;
	    %end;

    /* Set output PDF name depending on whether TOC is required or not */
    %if "&Create_TOC." eq "N" %then 
	    %do;
	        %let outpdf_bm = &OUTPDF.;
	    %end;
    %else 
	    %do;
	        %let outpdf_bm = &OUTPDF._bm;
	        %let inpdf_toc = &OUTPDF._bm;
	        %let outpdf_toc = &OUTPDF.;
	    %end;

    /* Importing BOOKMARK Excel sheet */
    %imp_xlsx(fname=&FILE_PATH.\Annotation.xlsx, sheet=BOOKMARK);

    /* Get current timestamp for unique file names */
    %let __ST = %sysfunc(datetime());
    data _null_;
        call symput('usertemp', sysget('TEMP'));
        start = compress(put(&__ST, is8601dt.), ,'kad');
        call symputx('__stime', start);  
    run;

    filename BMFILE "&USERTEMP.\BMARK_&__stime..txt";

    /* Create bookmark data */
    data _null_;
        file BMFILE notitles lrecl=1000 mod;
        set BOOKMARK;
            __BMTITLE = strip(BMTEXT);
            __BMLEVEL = strip(put(BMLEVEL, best.));
            __BMPAGE  = strip(put(BMPAGE, best.));
            put "BookmarkBegin";
            put "BookmarkTitle: " __BMTITLE;
            put "BookmarkLevel: " __BMLEVEL;
            put "BookmarkPageNumber: " __BMPAGE;
    run;

    /* Apply bookmarks to the input PDF using PDFtk */
    filename wzpipe pipe "pdftk ""&FILE_PATH.\&INPDF..pdf"" update_info ""&USERTEMP.\BMARK_&__stime..txt"" output ""&USERTEMP.\&INPDF..pdf"" ";

    data _null_;
        infile wzpipe length=l;
        input line $varying500. l;
    run;

    /* Move the output PDF with bookmarks */
    filename wzpipe pipe "move /Y ""&USERTEMP.\&INPDF..pdf"" ""&FILE_PATH.\&outpdf_bm..pdf"" ";

    data _null_;
        infile wzpipe length=l;
        input line $varying500. l;
    run;

    /* Create TOC if required */
    %if "&Create_TOC." eq "Y" %then 
    	%do;
        
	        /* Call Python script to create TOC */
	        x "python ""&FILE_PATH.\create_toc_hypl.py"" ""&FILE_PATH.\&inpdf_toc..pdf"" ""&FILE_PATH.\&outpdf_toc..pdf"" &TOC_SIZE.";
	        
	        /* Clean up temporary bookmark PDF */
	        x "del /f /q ""&FILE_PATH.\&OUTPDF_BM..pdf""";
	        
    	%end;

%mend add_bookmarks;

%add_bookmarks(INPDF=blankCRF_Positioned, Create_TOC=Y,TOC_SIZE=10);
