\name{read.gnumeric.range}
\title{Read data from a range of a gnumeric sheet}
\alias{read.gnumeric.range}
\description{
  
  Read data from a given cell range of a gnumeric (or other type
  supported by \code{ssconvert}) file to a data frame.

  By default drops no rows or columns, but requires at least bottom left
  corner of requested gnumeric cell range to be provided. Otherwise
  identical to \code{\link{read.gnumeric.sheet}}
  
  
}
\usage{
  df<-read.gnumeric.range(file,       
           head=FALSE, 
           sheet.name='Sheet1', 
           top.left='A1',
           bottom.right, 
           drop.empty.rows="none", 
           drop.empty.columns="none",
           colnames.as.sheet=FALSE,
           rownames.as.sheet=colnames.as.sheet,
           quiet=TRUE,
           ... 
           );
}

\seealso{\code{\link{read.gnumeric.sheet}}.}

\keyword{file}

