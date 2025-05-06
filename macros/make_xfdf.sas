/*****************************************************************************************************/
/* Macro Name     : make_xfdf                                                                        */
/*                                                                                                   */
/* Purpose        : Generates an XFDF file with annotations on each page, adhering to FDA guidelines.*/
/*                  Supports control of font size, font color, font type, box color, and border      */
/*                  color per domain, based on input from an Excel file.                             */
/*                                                                                                   */
/* Author         : Manivannan Mathialagan                                                           */
/* Created On     : 11-Mar-2022                                                                      */
/*                                                                                                   */
/* Parameters                                                                                        */
/*   - anno_location (Optional) : Path to the Excel file and blank CRF PDF. If not provided,         */
/*                                defaults to the directory where the macro is stored.               */
/*                                                                                                   */
/*   - page_orient   (Optional) : Orientation of PDF pages. Accepted values: Portrait or Landscape.  */
/*                                Default is Portrait.                                               */
/*                                                                                                   */
/*   - font_sz       (Optional) : Font size used for annotations. Accepts numeric values (e.g., 8, 9,*/
/*                                10). Default is 10.                                                */
/*                                                                                                   */
/*   - debug         (Optional) : Controls retention of intermediate datasets. Y = keep, N = delete. */
/*                                Default is N.                                                      */
/*                                                                                                   */
/* Example Usage                                                                                     */
/*   %make_xfdf(font_sz=12, debug=Y);                                                                */
/*                                                                                                   */
/* Notes                                                                                             */
/*   - The macro requires an Excel file named 'Annotation.xlsx' in the specified location.           */
/*   - The Excel file must contain the following two sheets:                                         */
/*       1) Annotation - containing annotation metadata per variable/domain.                         */
/*       2) Length     - character-level length values for text layout estimation.                   */
/*****************************************************************************************************/

%macro make_xfdf(anno_location=,
                 page_orient=Portrait,
                 font_sz=10,
                 debug=N);

/* If anno_location not provided, use script's directory */
%if "&anno_location." eq "" %then 
    %do;
        %let fullpath = %sysget(SAS_EXECFILEPATH);
        %let anno_location = %substr(&fullpath, 1, %length(&fullpath) - %length(%scan(&fullpath, -1, '\')) - 1);
        %put Full path: &fullpath;
        %put anno_location: &anno_location;
    %end;

/* Step 1: Import Excel */
%imp_xlsx(fname=&anno_location.\Annotation.xlsx);
%let font_sz_d = %eval(12 * &font_sz. / 10);

/* Define domain colors */
%let COLOR_ONE   = '#BFFFFF';
%let COLOR_TWO   = '#FFFF96';
%let COLOR_THREE = '#96FF96';
%let COLOR_FOUR  = '#FFBE9B';
%let COLOR_FIVE  = '#BFAAFF';

/* Filter and clean imported data */
data ANNO; 
    set ANNOTATION; 
    if not missing(compress(PAGENO,,'c'));
run;

data LEN;  
    set LENGTH;
run;

/* Step 2: Validate expected variables in input */
proc sql noprint;
    select upcase(NAME) into: varlist separated by ' '
    from SASHELP.VCOLUMN where LIBNAME='WORK' and MEMNAME='ANNO';
quit;

%macro misscheck(key=);
    %if %sysfunc(find(&varlist, &key)) <= 0 %then
        %put WARNING: Column &key missing from Excel input!;
%mend misscheck;

%misscheck(key=DOMAIN);
%misscheck(key=NAME);
%misscheck(key=PAGENO);
%misscheck(key=ANNOTATION);
%misscheck(key=ASSIGNEDFIELD);

/* Step 3: Process annotation data */
proc sql; create table SHELL (POSITION char(100), ORIENTATION char(100)); quit;

data ANNO2;
    set SHELL ANNO;
    NAME        = compress(NAME,,'c');
    ANNOTATION  = compress(ANNOTATION,,'c');
    DOMAIN      = compress(DOMAIN,,'c');
    DOMAIN1     = DOMAIN;
    PAGENO_     = compress(PAGENO,,'c');
    POSITION    = compress(POSITION,,'c');
    ORIENTATION = compress(ORIENTATION,,'c');
    if missing(ORIENTATION) then ORIENTATION = "&page_orient";
    if not missing(ANNOTATION);
    drop PAGENO;
run;

/* Expand multiple page numbers */
data ANNO3;
    set ANNO2;
    do i = 1 to 20;
        PAGENO = compress(scan(PAGENO_, i, ','));
        if not missing(PAGENO) then PAGENO = strip(put(input(PAGENO, best.) - 1, best.));
        output;
    end;
run;

data ANNO3;
    set ANNO3(where=(not missing(compress(PAGENO, ' ', 'c'))));
    ORD = _n_;
run;

proc sort data=ANNO3; 
    by PAGENO DOMAIN1 ORD; 
run;

/* Step 4: Assign text lengths and box positions */
data ANNO4;
    length CHAR $1.;
    retain ORD_DOMAIN;
    set ANNO3;
    by PAGENO DOMAIN1 ORD;

    AUTHOR  = DOMAIN;
    SUBJECT = NAME;
    DOMAIN_ = lag(DOMAIN1);

    if first.PAGENO then ORD_DOMAIN = 1;
    else if DOMAIN1 ne DOMAIN_ then ORD_DOMAIN + 1;

    if first.DOMAIN1 then ORD_VAR = 1;
    else if missing(POSITION) then ORD_VAR + 1;

    call missing(CHAR, LEN);
    declare hash h(dataset: 'work.len', ordered: 'ascending');
    h.definekey('char');
    h.definedata('len');
    h.definedone();

    tot = 0;
    do i = 1 to length(ANNOTATION);
        rc = h.find(key: substr(ANNOTATION, i, 1));
        tot + len;
    end;

    select (ORD_DOMAIN);
        when (1) BACKCOLOR = &COLOR_ONE;
        when (2) BACKCOLOR = &COLOR_TWO;
        when (3) BACKCOLOR = &COLOR_THREE;
        when (4) BACKCOLOR = &COLOR_FOUR;
        when (5) BACKCOLOR = &COLOR_FIVE;
        otherwise;
    end;

    n1 = _n_;
    n2 = _n_ + 10000;

    if ORIENTATION = 'Portrait' then do;
        x1 = 0 + (ORD_VAR - 1) * &font_sz. + 20 * (ORD_DOMAIN - 1);
        x2 = x1 + tot + 5;
        y1 = 500 - (ORD_VAR - 1) * &font_sz.;
        y2 = y1 + 15;
        x1_d = 0 + 20 * (ORD_DOMAIN - 1);
        x2_d = x1_d + tot * (&font_sz_d. / &font_sz.) + 5 * (&font_sz_d. / &font_sz.);
        y1_d = 510;
        y2_d = y1_d + 15 * (&font_sz_d. / &font_sz.);
    end;
    else if ORIENTATION = 'Landscape' then do;
        y1 = 0 + (ORD_VAR - 1) * &font_sz. + 20 * (ORD_DOMAIN - 1);
        y2 = y1 + tot + 5;
        x1 = 30 + (ORD_VAR - 1) * &font_sz.;
        x2 = x1 - 15;
        y1_d = 0 + 20 * (ORD_DOMAIN - 1);
        y2_d = y1_d + tot * (&font_sz_d. / &font_sz.) + 5 * (&font_sz_d. / &font_sz.);
        x1_d = 20;
        x2_d = x1_d - 15 * (&font_sz_d. / &font_sz.);
    end;
run;

/* Step 5: Write XFDF output */
data _NULL_;
    file "&anno_location.\xfdf_not_positioned.xfdf";
    set ANNO4 end=end;
    
    by PAGENO ORD_DOMAIN;

    if _n_=1 then
        do;
            put '<?xml version="1.0" encoding="UTF-8"?>';
            put '<xfdf xmlns="http://ns.adobe.com/xfdf/" xml:space="preserve"><annots>';
        end;

    if missing(NAME) and not index(upcase(ANNOTATION),'NOT SUBMITTED') then
        do;
            put '<freetext';
            put ' flags="print"';
            put ' color="' backcolor '"';
            put ' name="' n2 '"';
            put ' page="' pageno '"';

            if ASSIGNEDFIELD eq 'Y' then
                do;
                    put ' dashes="3.000000,3.000000" style="dash"';
                end;
            else 
                do;
                    put ' style="solid"';
                end;
            if ORIENTATION = 'Landscape' then
                do;
                    put ' rotation = "90"';
                end;

            if missing(POSITION) then
                do;
                    put ' rect="' x1_d ',' y1_d ',' x2_d ',' y2_d '"';
                end;

            if not missing(POSITION) then
                do;
                    put ' rect="' position '"';
                end;

            put ' subject="' subject '"';
            put ' title="' author '">';
            put '<contents-richtext>';
            put '<body xmlns="http://www.w3.org/1999/xhtml" xmlns:xfa="http://www.xfa.org/schema/xfa-data/1.0/"';
            put ' xfa:APIVersion="Acrobat:11.0.0" xfa:spec="2.0.2" ';
            put " style='font-size:&font_sz_d. pt;text-align:left;color:#000000;font-weight:bold;";
            put " font-style:normal;font-family:Arial;font-stretch:normal'";
            put '><p dir="ltr">' annotation '</p';
            put '></body></contents-richtext></freetext>';
        end;
    else
        do;
            put '<freetext';
            put ' flags="print"';
            put ' color="' backcolor '"';
            put ' name="' n1 '"';
            put ' page="' pageno '"';
            if  ASSIGNEDFIELD eq 'Y' then
                do;
                    put ' dashes="3.000000,3.000000" style="dash"';
                end;
            else 
                do;
                    put ' style="solid"';
                end;
            if ORIENTATION = 'Landscape' then
                do;
                    put ' rotation = "90"';
                end;

            if missing(POSITION) then
                do;
                    put ' rect="' x1 ',' y1 ',' x2 ',' y2 '"';
                end;

            if not missing(POSITION) then
                do;
                    put ' rect="' position '"';
                end;

            put ' subject="' subject '"';
            put ' title="' author '">';
            put '<contents-richtext>';
            put '<body xmlns="http://www.w3.org/1999/xhtml" xmlns:xfa="http://www.xfa.org/schema/xfa-data/1.0/"';
            put ' xfa:APIVersion="Acrobat:11.0.0" xfa:spec="2.0.2" ';
            put " style='font-size:&font_sz. pt;text-align:left;color:#000000;font-weight:normal;";
            put " font-style:normal;font-family:Arial;font-stretch:normal'";
            put '><p dir="ltr">' annotation '</p';
            put '></body></contents-richtext></freetext>';
        end;

    if end then
        do;
            put '</annots>';
            put '</xfdf>';
        end;
run;

/* Step 6: Cleanup */
%if %upcase(&debug.) ne Y %then %do;
    proc datasets lib=work nolist; delete ANNO ANNO2 ANNO3 ANNO4 SHELL LEN; quit;
%end;

%mend make_xfdf;

/* Example execution */
%make_xfdf();
