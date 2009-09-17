\name{read.gnumeric.sheets}

\title{Read each sheet from a gnumeric spreadsheet file}

\alias{read.gnumeric.sheets}

\description{
  
  Read data from each non-empty sheet of a gnumeric file to a list of
  data.frames.
  
}
\usage{

df.list <- read.gnumeric.sheets(file,
                                head=FALSE,
                                drop.empty.rows="none",
                                drop.empty.columns="none",
                                colnames.as.sheet=FALSE,
                                rownames.as.sheet=colnames.as.sheet,
                                quiet=TRUE,
                                ...  
                                );
}

\arguments{
  
  \item{file}{Name of gnumeric file to read from.}
  
  \item{head, drop.empty.rows, drop.empty.columns}{}
  \item{colnames.as.sheet,rownames.as.sheet, quiet, ...}{
    These are passed to \code{\link{read.gnumeric.sheet}}
  }
  
}

\value{
  A list of data frames. 
}

\seealso{
  
  \code{\link{read.gnumeric.sheet.info}} to list sheet names and
  sizes.
  
  \code{\link{read.gnumeric.range}} to read an exact cell range of a
  single sheet.
  
}

%%\details{}

\examples{
## Read all sheets from 'file.gnumeric'
\dontrun{
df.list <- read.gnumeric.sheets( file="file.gnumeric" );

df1 <- df.list['Sheet1'];

df.list <- read.gnumeric.sheets( file="file.gnumeric",
                                 head=TRUE );

}

}

\keyword{IO}
\keyword{file}

