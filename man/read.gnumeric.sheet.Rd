\name{read.gnumeric.sheet}

\title{Read data from a gnumeric (or MS Excel, Openoffice Calc, Xbase,
Quatro Pro, Paradox, HTML, etc) spreadsheet or database file using
ssconvert from the gnumeric distribution}

\alias{read.gnumeric.sheet}

\description{
  
  Read data from a sheet of a gnumeric (or other common spreadsheet or
  database) file to a data.frame.
  
  Requires an external program, \samp{ssconvert} (normally installed
  with gnumeric (\url{http://projects.gnome.org/gnumeric/}) in \samp{PATH}.
  
  Calls \samp{ssconvert} to convert the input to CSV.
  \samp{ssconvert} can read several file formats
  (see Details below).

  Note: During conversion to CSV \samp{ssconvert} also evaluates formulas
  (e.g. \samp{=sum(A1:A3)}) in cells, and emits the result instead of
  the formula.
}
\usage{
  df<-read.gnumeric.sheet(file,
                          head=FALSE,
                          sheet.name='Sheet1',
                          top.left='A1',
                          bottom.right='IV65536',
                          drop.empty.rows="bottom", 
                          drop.empty.columns="right",
                          colnames.as.sheet=FALSE,
                          rownames.as.sheet=colnames.as.sheet,
                          quiet=TRUE,
                          LANG='C',
                          import.encoding=NA,
                          ...
                         )

}

\arguments{
  
  \item{file}{Name of gnumeric file (or other file type readable by
    gnumeric) to read from.

    This may also be an URL, i.e. like\cr
    \code{'http://example.com/path/file.gnumeric'}
  
  }
  
  \item{head}{When TRUE, use first row of requested gnumeric sheet
    range as column names in the resulting \code{data.frame}}
  
  \item{sheet.name}{Name of sheet as appears in gnumeric.\cr
    Sheet names containing space character do not work.
  
    \code{sheet.name=NA} Omits sheet name from the ssconvert command line.\cr
    For gnumeric files this will read the sheet that was 'current' in
      gnumeric when the file was saved. 
  }
  
  \item{top.left}{Top left corner of requested gnumeric sheet
    range, e.g. \code{'A1'}}
  
  \item{bottom.right}{Bottom right corner of requested gnumeric sheet
    range.\cr
    The default for \code{read.gnumeric.sheet} is \code{'IV65536'}: this
    causes a lot of unused lines to be printed by \samp{ssconvert} then
    parsed by \code{read.csv}, thus you might want to override it to
    speed up reading.

    Use \code{\link{read.gnumeric.sheet.info}} to read actual
    bottom.right cell name from a gnumeric file (but not other formats).
  }

  \item{drop.empty.rows}{One of
    \code{c('none','top','bottom','both','all')}.

    \code{'all'}  drops all empty lines from the requested range, even
    those that are between two non-empty rows.
    \code{'both'} drops empty lines below the last non-empty row and
    above the first non-empty.
    \code{'top'}, \code{'bottom'} and \code{'none'} as you would expect.
  }

  \item{drop.empty.columns}{ One of
    \code{c('none','left','right','both','all')}\cr
    Similar to
    \code{drop.empty.rows}, but for columns.
  }

  \item{colnames.as.sheet}{Rename columns to \code{'A', 'B', 'C', ...} to
    have names corresponding to gnumeric column names.}

  \item{rownames.as.sheet}{Rename rows to \code{'1', '2', '3', ...} to have
    names corresponding to gnumeric row indices. Note: this means
    \code{df['1',]}, not \code{df[1,]} in the result (rownames are
    strings, not integers). Note: when deciding row names only
    \code{top.left} and \code{head} are accounted for, but not
    e.g. \code{skip} (which may be passed to \code{read.csv} via \code{...}).
  }
  
  \item{quiet}{When TRUE, do not print command executed, and (on unix
    platforms) also redirect stderr of the external program
    \samp{ssconvert} to /dev/null}
  
  \item{LANG}{ Under unix, passed to \code{ssconvert} in the environment
    variable 'LANG'. The default value (\code{'C'}) is intended to avoid
    using decimal comma in the emitted CSV file. }

  \item{import.encoding}{If not NA, passed to \code{ssconvert} as its
    \code{ --import-encoding } parameter.
    % One might also need to pass encoding='XXX' in ... to read.csv if
    % encoding issues arise.
  }
  
  \item{...}{Extra arguments, passed to \code{\link{read.csv}}}
}


\seealso{

  \code{\link{read.gnumeric.range}} for a variant with default
  arguments more suited for reading an exact cell range of a sheet.
  

  \code{\link{read.gnumeric.sheet.info}} to read actual
  bottom.right cell name from a gnumeric file (but not other formats).
  
  \code{\link{read.gnumeric.sheets}} to read all sheets
  from a gnumeric file (but not other formats).
  
  
  \code{\link[gdata]{read.xls}} for reading Microsoft Excel files
  (possibly from a \samp{http://} URL)

  \code{\link{read.DIF}} for reading Data Interchange Format (DIF)
  files.

  \code{\link[foreign]{read.dbf}} for Xbase (.dbf) files.
}

\details{

  Data from the gnumeric file is dumped as .csv using the \samp{ssconvert}
  program provided with gnumeric.

  \samp{ssconvert} supports several input formats, thus the input file
  does not have to be a gnumeric file. The formats supported may be
  listed with
  \preformatted{
    ssconvert --list-importers
  }
  from a shell prompt.

  For me this prints (with ssconvert version '1.8.4')
  \preformatted{
ID                           | Description
Gnumeric_xbase:xbase         | Xbase (*.dbf) file format
Gnumeric_Excel:excel         | MS Excel (tm) (*.xls)
Gnumeric_Excel:xlsx          | MS Excel (tm) 2007
Gnumeric_html:html           | HTML (*.html, *.htm)
Gnumeric_oleo:oleo           | GNU Oleo (*.oleo)
Gnumeric_applix:applix       | Applix (*.as)
Gnumeric_QPro:qpro           | Quattro Pro (*.wb1, *.wb2, *.wb3)
Gnumeric_paradox:paradox     | Paradox database or
                             | primary index file
Gnumeric_sc:sc               | SC/xspread
Gnumeric_XmlIO:sax           | Gnumeric XML (*.gnumeric)
Gnumeric_lotus:lotus         | Lotus 123 (*.wk1, *.wks, *.123)
Gnumeric_XmlIO:dom           | Gnumeric XML (*.gnumeric) Old
                             |   slow importer
Gnumeric_dif:dif             | Data Interchange Format (*.dif)
Gnumeric_Excel:excel_xml     | MS Excel (tm) 2003 SpreadsheetML
Gnumeric_OpenCalc:openoffice | Open/Star Calc (*.sxc, *.ods)
Gnumeric_plan_perfect:pln    | Plan Perfect Format (PLN) import
Gnumeric_sylk:sylk           | MultiPlan (SYLK)
Gnumeric_mps:mps             | Linear and integer program (*.mps)
                             |   file format
Gnumeric_stf:stf_csvtab      | Comma or tab separated
			     |   values (CSV/TSV)
Gnumeric_stf:stf_assistant   | Text import (configurable)
  }

  But the actual list may depend on
  which import plugins are installed for gnumeric.
  
  \tabular{lll}{
    Format \tab Source \tab Status\cr
  .gnumeric  \tab gnumeric \tab works\cr
  .xls  \tab gnumeric \tab works\cr
  .html \tab gnumeric  \samp{[Save as / HTML 4.0]}  \tab works\cr
  .html \tab Openoffice Calc \samp{[Save as/HTML Document]} \tab works\cr
  .ods \tab  Openoffice Calc \tab might work[1]\cr
  Other formats \tab \tab not tested
    }
  
    [1] during .ods import ssconvert emits some messages on stdout
    (corrupts output). Worked around on unix with \samp{| grep ,}
    on the output. (The problem has been corrected in libgsf
    on 2009-08-20, thus the workaround is expected to be removable
    by 2010)
  
%  Note: one could also write several file formats by feeding CSV to
%  ssconvert (but that is not implemented here)
}

\examples{
## Read all data from 'Sheet1'
\dontrun{
df <- read.gnumeric.sheet( file="file.gnumeric" );

df <- read.gnumeric.sheet( file="file.gnumeric",
                           sheet.name='Sheet1' );



## Read from Excel sheet named 'Sheet3' the range C3:D50,
## rename columns to 'C' and 'D', rows to '3' ... '50',
## then drop all empty rows.
## 
df<-read.gnumeric.sheet( "file.xls",
                          sheet.name='Sheet3',
                          top.left='C3',
                          bottom.right='D50',
                          drop.empty.rows="all", 
                          drop.empty.columns="none",
                          colnames.as.sheet=TRUE
                         )


## Read from "file.gnumeric", 'Sheet1' data in 'A1:E100',
## Use first row (of selected range) as column names.
## Drop empty rows and columns from bottom and right.
df<-read.gnumeric.sheet("file.gnumeric", head=TRUE,
                         bottom.right='E100')


## Why does it not work? Set quiet=FALSE to see
## the command executed (and on unix, diagnostic
## messages from ssconvert).
df<-read.gnumeric.sheet( "file.ods", quiet=FALSE )
}

}

\keyword{IO}
\keyword{file}

