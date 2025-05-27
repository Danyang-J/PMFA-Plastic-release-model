excel.file <- "Sample_FeedData.xlsx"

# source needed functions
source("Code/functions.needed.R")
source("Code/functions.needed.analysis.R")

library(openxlsx)
library("trapezoid")

Materials <- c("LDPE", "HDPE", "PP", "PS", "EPS", "PVC", "PET")
Systems <- "CH"
SIM <- 10^4

message(paste0("\n\n",format(Sys.time(), "%H:%M:%S"), " Module 1 definition"))
source("Code/1-InputFormatting.R")
message(paste0("\n\n",format(Sys.time(), "%H:%M:%S"), " Module 1 done !"))

message(paste0("\n\n",format(Sys.time(), "%H:%M:%S"), " Module 2 definition"))
source("Code/2-Merging.R")
message(paste0("\n\n",format(Sys.time(), "%H:%M:%S"), " Module 2 done !"))

source("Code/3-CalculationScript.R")

# Plot results
source("Code/Graph-EmissionsByProd.R")

