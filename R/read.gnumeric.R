
## sheet.name defaults to 'Sheet1' instead of NA, which would get the
## 'current sheet' from the .gnumeric file.
##
## Requirements:
##
## ssconvert (spreadsheet converter program, comes with gnumeric)
##
##            We use its undocumented flag:
##            --export-range='Sheet1!A1:Z100' or
##            --export-range='A1:Z100' (read whichever sheet was 'current'
##            in gnumeric when the file was saved)
##
## quiet=TRUE uses '2>/dev/null' redirection (when .Platform$OS.type=="unix" )
##
##
## ??? Character encoding issues? gnumeric appears to use utf-8 (at
## least under an utf-8 locale)
##
read.gnumeric.sheet <-
  function(file,       ## 'filename.gnumeric'
           head=FALSE, ## as in read.csv, but first row seen by
                       ## read.csv is decided by row in top.left
           sheet.name='Sheet1', ## name of the sheet as appears in
                                ## gnumeric
           top.left='A1',       ## top left cell to request from
                                ## ssconvert (gnumeric utility)
           bottom.right='IV65536', ## bottom right cell. The default
                                   ## reads all rows, which is
                                   ## slow. Speed up by giving a more
                                   ## accurate upper bound, or
                                   ## accurate value.
           drop.empty.rows="bottom", ## Drop rows containing only NAs
                                     ## and empty strings from result
           drop.empty.columns="right",## Drop columns containing only
                                      ## NAs and empty strings from
                                      ## result
           colnames.as.sheet=FALSE,   ## When TRUE, names(result) are
                                      ## set to those of
                                      ## gnumeric. (Overwrites names
                                      ## from read.csv)
           rownames.as.sheet=colnames.as.sheet, ## When TRUE, rownames
                                                ## are set to gnumeric
                                                ## row
                                                ## indices. (head=TRUE
                                                ## is taken into
                                                ## account, but e.g. skip!=0 not)
           quiet=TRUE,                ## Redirect stderr of ssconvert
                                      ## to /dev/null and do not print
                                      ## command executed
           ... ## passed to read.csv
           )
{

  ## 'all': drop even those between row (or columns) contining data
  if ( ! drop.empty.rows %in% c('none','top','bottom','both', 'all' ) ){
    stop( "drop.empty.rows is not in c('none','top','bottom','both', 'all' )" )
  }
  if ( ! drop.empty.columns %in% c('none','left','right','both', 'all' ) ){
    stop( "drop.empty.columns is not in c('none','left','right','both', 'all' )" )
  }



  ### build command
  SHEET='';
  if ( !is.na(sheet.name) ){
    SHEET=paste(sheet.name, '!', sep='' );
  }


  ssconvert = "ssconvert";
  if ( Sys.which( ssconvert ) == "" ){
    stop("Required program '",ssconvert,"' not found." );
  }

  ## --export-range needed because I know of no other way to select the
  ## sheet. This in turn forces to also provide top.left and
  ## bottom.right, even when we just want 'all the sheet'
  cmd <- paste(ssconvert,
    " --export-type=Gnumeric_stf:stf_csv ",
    " --export-range='", SHEET , top.left ,":", bottom.right,"' ",
    "'", file, "'",
    " fd://1 ", sep='');


  
  if ( ! quiet ){
    cat(cmd,"\n")
  } else {
    if ( .Platform$OS.type == "unix" ){
      cmd = paste( cmd, " 2> /dev/null" ); ## unix
    } else {
      ## ( .Platform$OS.type == "windows" )
      ## ??? 2>NIL:  ??? or similar? 
    }
  }

  ### read data
  x=read.csv( pipe( cmd  ) , head=head, ... )

  ### optionally rename columns and rows to correspond to gnumeric
  ### cell names.
  if ( colnames.as.sheet || rownames.as.sheet ){
    ABC=LETTERS;
    ## COLNAMES: A .. AA, AB, .. IV
    COLNAMES= as.vector( t(outer(c('',ABC[1:9]), ABC, paste, sep='')))[1:256]; 

    left=''
    i=1;
    while ( substr( top.left,i,i) %in% ABC ){
      i=i+1;
    }
    left = substr( top.left, 1, (i-1) )
    top.str = substr( top.left, i, nchar(top.left) );
    top=as.integer( top.str );

    if ( colnames.as.sheet ){
      left.index = match( left, COLNAMES, nomatch=0 )
      stopifnot( left.index > 0 )
      names(x)=COLNAMES[left.index:(left.index+length(x)-1)];
    }

    if ( rownames.as.sheet ){

      if ( head ){
        top = top+1;
      }
      
      rownames(x) = top:(top+length(x[[1]])-1)
    }
  }

  ### optionally drop empty columns and rows 
  if ( drop.empty.columns != "none" || drop.empty.rows!="none" ){
    ## drop empty columns and rows
    last.col=length(x);
    last.row=length(x[[1]]);

    is.empty <- function(x){ all( is.na(x) | x=='' )  }
    m=as.matrix(x);

    i=1:last.row; ## indices of rows to keep

    if ( drop.empty.rows != "none" ){
      bi=!apply(m,1,is.empty);
      i=i[bi]; ## indices of non-empty rows
      
      if ( drop.empty.rows == "bottom" ){
        ## extend i to include all rows from row 1 to last non-empty
        ## (includes empty rows between non-empty ones as well)
        i = 1:i[length(i)]; 
      } else if ( drop.empty.rows == "top" ){
        i = i[1]:last.row
      } else if ( drop.empty.rows == "both" ){
        i = i[1]:i[length(i)]
      }
    }


    j=1:last.col;
    if ( drop.empty.columns  != "none"){
      bj=!apply(m,2,is.empty)
      j=j[bj]
      if ( drop.empty.columns == "right" ){
        j = 1:j[length(j)]
      } else if ( drop.empty.columns == "left" ){
        j = j[1]:last.col
      } else if ( drop.empty.columns == "both" ){
        j = j[1]:j[length(j)]
      }
    }
    x=x[i,j]
  }
  
  x
}



## Like read.gnumeric.sheet, but bottom.right is mandatory and we drop
## no rows or columns by default.
read.gnumeric.range <-
  function(file,       
           head=FALSE, 
           sheet.name='Sheet1', 
           top.left='A1',
           bottom.right, 
           drop.empty.rows="none", 
           drop.empty.columns="none",
           colnames.as.sheet=FALSE,
           rownames.as.sheet=colnames.as.sheet,
           quiet=TRUE,
           ... ## passed to read.csv
           )
{
  read.gnumeric.sheet(file,
                      head,
                      sheet.name,
                      top.left,
                      bottom.right,
                      drop.empty.rows,
                      drop.empty.columns,
                      colnames.as.sheet, 
                      rownames.as.sheet,
                      quiet,
                      ... )
}


