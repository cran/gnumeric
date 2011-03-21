
source('read.gnumeric.R');

x=read.gnumeric.sheet('Book1.gnumeric', sheet='Sheet1');
x$V4.a = gnumeric.as.date(x$V4);
x$V5.a = gnumeric.as.date(x$V5);
x
