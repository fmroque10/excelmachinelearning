


  
#install.packages("scatterplot3d")   
library("scatterplot3d")

rng <- EXCEL$Application$get_Range( "B2:D35" )
X <- rng$get_Value()


scatterplot3d(x=X[,1],y=X[,2],z=X[,3])
