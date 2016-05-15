rm(list=ls(all=TRUE))

source("trpv1.R")

global_noise_index = 120

# Process files
trpv_file("04-19-16_slip1.xlsx",11, global_noise_index)
trpv_file("04-13-16_slip1.xlsx",11, global_noise_index)