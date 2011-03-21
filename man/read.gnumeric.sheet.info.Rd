\name{read.gnumeric.sheet.info}

\title{Read names and sizes of sheets from a gnumeric spreadsheet file}

\alias{read.gnumeric.sheet.info}

\description{
  
  Read \code{sheet.name}, \code{width} and \code{height} of sheets of a
  gnumeric file to a \code{data.frame}. Also constructs cell name for the
  \code{bottom.right} cells, these can be passed to
  \code{\link{read.gnumeric.sheet}} or \code{\link{read.gnumeric.range}}
  
}
\usage{
  read.gnumeric.sheet.info(file);
}

\arguments{
  \item{file}{Name of gnumeric file to read from.}
}

\value{
  A \code{data.frame} with colummns
  \tabular{lll}{
    Column name       \tab Type   \tab Note \cr
    \code{sheet.name} \tab string \tab \cr
    \code{width}      \tab integer \tab Number of columns \cr
    \code{height}     \tab integer \tab Number of rows \cr
    \code{bottom.right} \tab string \tab Bottom right cell name or \cr
                        \tab        \tab \code{NA} if sheet is empty \cr
  }
}

\examples{
## Read names and sizes of sheets from 'file.gnumeric'
\dontrun{
df <- read.gnumeric.sheet.info( file="file.gnumeric" );
}
}

\keyword{IO}
\keyword{file}

