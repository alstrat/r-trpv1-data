rm(list=ls(all=TRUE))

source("trpv1.cfg")

library(xlsx)
library(gridExtra)
library(ggplot2)
library(reshape2)


# logging helper
logVar <- function(string, var) {
    print(paste(string, " ", var))
}


########################### Read XLSX ###########################
get_first_data_length <- function(data_sheet){
    options(warn=-1)
    first_row <- readColumns(data_sheet, startRow = 2, endRow = 2, startColumn = 1, endColumn = 100, header=FALSE)
    first_blank_index <- min(which(first_row == ""))
    options(warn=0)
    return((first_blank_index - 1))
}


get_sheet_first_data <- function(file_name, sheet_number, dataType){

    file_wb <- loadWorkbook(file_name)
    file_sheets <- getSheets(file_wb)
    data_sheet <- file_sheets[[sheet_number]]
    first_data_length <- get_first_data_length(data_sheet)
    raw <- read.xlsx2(file_name, 
            sheetIndex = sheet_number, 
            colIndex = (1:first_data_length),
                  colClasses=rep(dataType, first_data_length))

    data <- data.matrix(raw[1:first_data_length])
    return(data)
}


# Process 1 sheet of input spreadsheet. TODO: rename
get_input_sheet_data <- function(file_name, sheet_number){
    raw_data <- get_sheet_first_data(file_name, sheet_number, "numeric")
    
    bg_index <- ncol(raw_data)
    bg <- raw_data[,bg_index]
    time <- raw_data[,1]
    trpv_data <- raw_data[,2:(bg_index-1)]
    
    processed_data <- apply(trpv_data, 2, process_trpv_col, bg=bg, time=time)

    return(processed_data)
}


########################### basic plot on matplot ###########################
plot_initial <- function(file_name, trpv1, aktph) {
   png_output_file_name = paste(file_name, "_out.png", sep="")
   
   logVar("    [START] Building plots into ", png_output_file_name)
   
    png(png_output_file_name,
        width     = 18,
        height    = 9,
        units     = "in",
        res       = 1200,
        pointsize = 4
    )
    par(c(1,2), xaxs = "i", yaxs = "i", cex.axis = 2, cex.lab  = 2)
    layout(matrix(c(1,2), 1, 2, byrow = TRUE), widths=c(1,1), heights=c(1,1))
    
    matplot(aktph, type = "l")
    m_names = names(aktph[1,])
    nn <- ncol(aktph)
    legend("bottomleft", legend = m_names,col=seq_len(nn),cex=2, fill=seq_len(nn))
    
    matplot(trpv1, type = "l")
    m_names = names(trpv1[1,])
    nn <- ncol(trpv1)
    legend("bottomleft", legend = m_names,col=seq_len(nn),cex=2, fill=seq_len(nn))
    dev.off()
    
    logVar("    [END] Building plots into ", png_output_file_name)
 }


########################### Input processing ###########################

# Process table. TODO: rename
process_trpv_col <- function(in_trpv_column, bg, time) {

    # Remove BG and 
    # trpv_column_no_bg_reduced <- vapply((in_trpv_column - bg), reduce_bg_noise, 0)
    trpv_column_no_bg_reduced <- (in_trpv_column - bg)
    trpv_column <- trpv_column_no_bg_reduced

    # Build linear regression model cell(t) on 1:noise_index
    y_lm <- trpv_column[1:noise_index] 
    x_lm <- time[1:noise_index]
    lin_reg <- lm(formula = y_lm ~ x_lm)
    
    # Dirty hack to subsract predicted regression(a(x) + intercep) but keep intercept value: cell - a(x) - intercep + intercept 
    intercept_lm <- summary(lin_reg)$coefficients[1,1]
    x_lm_predict <- time
    new_x_lm <- data.frame("x_lm"=x_lm_predict)
    new_y_lm <- predict(lin_reg, newdata=new_x_lm)
    cleansed_trpv_column <- trpv_column - new_y_lm + rep(intercept_lm, (length(time)))
    
    # correction on mean and standard deviation of baseline period
    # take all and divide by mean from baseline period = x/mean(x[1:noise_index]), 
    # standard deviation currently not used
    mean_noaction <- mean(cleansed_trpv_column[1:noise_index])
    # sd_noaction <- sd(cleansed_trpv_column[1:120]) 
    # normalized_trpv <- vapply(cleansed_trpv_column, function(x) (abs(x) - sd_noaction)/mean_noaction * sign(x), 0)
    normalized_trpv <- vapply(cleansed_trpv_column, function(x) x/mean_noaction, 0)
    
    # normalized_trpv <- vapply(cleansed_trpv_column, function(x) (x - min(x))/(max(x) - min(x)), 0)
    
    return(normalized_trpv)
}


# Process input excel document with 2 tabs
# tab 1 - aktph; tab 2 - trpv1
# returns name of excel file with 2 tables and optional plot in png as sideeffect
#TODO: rename
process_trpv1_input_file <- function(in_file_name, draw_plot_flag) {

    # build output files basename based on input filename
    #   add extentions and suffixes
    #   output in out/ subfolder
    file_name <- paste("in/", in_file_name, sep="")
    out_base_name <- paste("out/", (sub(".xlsx", "", in_file_name)), sep="")
    
    # logging steps to console for progress monitoring
    logVar("[START] Processing", file_name)
    logVar("    index where noise ends", noise_index)
    
    # process both tabs as function above: remove bg, linear regression of baseline, bring to 1
    aktph <- get_input_sheet_data(file_name, AKTPH_INDEX)
    trpv1 <- get_input_sheet_data(file_name, TRPV_INDEX)

    # optional, draw plot into PNG
    if (draw_plot_flag == 1 ) {
        plot_initial(out_base_name, trpv1, aktph)
    }
   
    # dump resulting tables into excel
    output_file_name = paste(out_base_name, "_out.xlsx", sep="")
    
    logVar("    [START] Dumping results into", output_file_name)
    
    write.xlsx(aktph, file=output_file_name, sheetName="processed_aktph")
    write.xlsx(trpv1, file=output_file_name, sheetName="processed_trpv1", append=TRUE)
    
    logVar("    [END] Dumping results into", output_file_name)
    logVar("[END] Processing", file_name)
    
    return(output_file_name)
}


################################# grouping helpers #########################################


################################# PLOTING #########################################
draw_plot_list_into_file <- function(plots, file_name) {

    png_output_file_name = paste(file_name, "_out.png", sep="")

    logVar("    [START] Building plots into ", png_output_file_name)

    pg <- marrangeGrob(plots, nrow=length(plots), ncol=1)
    ggsave(file=png_output_file_name, pg, dpi = 300, width = 30, height = 24, units = "in")
    logVar("    [END] Building plots into ", png_output_file_name)
}

# plot drawing helper
plot_data <- function(data, title) {
    d <- melt(data, id.vars="t")
    # Everything on the same plot
    data_plot <- ggplot(d, aes(Var1,value, col=Var2))+ 
        geom_line(size=1.2)+
        labs(x="t", y="cells", title=title) +
        coord_cartesian(xlim=c(100, 360), ylim=c(-1, 3))+
        scale_x_continuous(breaks=seq(0, 500, 20))+
        scale_y_continuous(breaks=seq(-1, 3, 0.2))
    return(data_plot)
}

plot_groups_independent <- function(grouped_data, file_name){
    # aktph_plot <- plot_data(grouped_data[[AKTPH_INDEX]], "aktph")
    # trpv_plot <- plot_data(grouped_data[[TRPV_INDEX]], "trpv")
    # trpv_aktph_plot <- plot_data(grouped_data[[TRPV_AKTPH_INDEX]], "trpv_aktph")
    # plots <- list(aktph_plot, trpv_plot, trpv_aktph_plot)
    
    data_names <- names(grouped_data)
    plots <- lapply(data_names, function(x) plot_data(grouped_data[x], x))
    
    draw_plot_list_into_file(plots, file_name)
}


# group determination logic
# test date metrics(max and mean) are compared to baseline with a threshold,
# defined as mean value of baseline sample with given bandwidth
#
# Group is HIGH if either mean or maximum of data sample is bigger them threshold upper bound;
# group is LOW if mean value of data sample is lesser them threshold lower bound;
# otherwise group is FLAT
group_against_base_aktph <- function(test_data, base){
    base_mean <- mean(base)
    THRESHOLD_UPPER_BOUND <- base_mean + THRESHOLD_BANWIDTH
    THRESHOLD_LOWER_BOUND <- base_mean - THRESHOLD_BANWIDTH
    
    data_mean <- mean(test_data)
    data_max <- max(test_data)
    
    if ((data_mean > THRESHOLD_UPPER_BOUND) | (data_max > THRESHOLD_UPPER_BOUND)) { 
        group <- "HIGH"
    } else if (data_mean < THRESHOLD_LOWER_BOUND) {
        group <- "LOW"
    } else {
        group <- "FLAT"
    }
    return(group)
}


# Get group label for a given data column
#   column elements from 1 to noise_index considered as baseline
#   column elements from (noise_index + 1) to (TEST_DATA_SIZE * noise_index) considered as sample to choose group
# TODO: refactor
get_data_group_aktph <- function(data) {
    
    TEST_DATA_MIN <- 140
    TEST_DATA_MAX <- 180
    
    base<- data[1: noise_index]
    # sample to determine group
    test_data <- data[TEST_DATA_MIN: TEST_DATA_MAX]
    # group determination
    group <- group_against_base_aktph(test_data, base)
    return(group)
}

group_against_base_trpv <- function(test_data, base){
    base_mean <- mean(base)
    THRESHOLD_UPPER_BOUND <- base_mean + THRESHOLD_BANWIDTH
    THRESHOLD_LOWER_BOUND <- base_mean - THRESHOLD_BANWIDTH
    
    data_mean <- mean(test_data)
    data_max <- max(test_data)
    
    if ( data_max > THRESHOLD_UPPER_BOUND) { 
        group <- "HIGH"
    } else if (data_mean < THRESHOLD_LOWER_BOUND) {
        group <- "LOW"
    } else {
        group <- "FLAT"
    }
    return(group)
}

get_data_group_trpv <- function(data) {
    
    TEST_DATA_MIN <- 120
    
    base<- data[1: noise_index]
    # sample to determine group
    test_data <- data[-(1:TEST_DATA_MIN)]
    # group determination
    group <- group_against_base_aktph(test_data, base)
    return(group)
}


# helper to add labels to column headers
append_headers <- function(data, appenders){

    original_colnames <- colnames(data)
    new_colnames <- paste(appenders, original_colnames, sep=".")
    colnames(data) <- new_colnames
    return(data)
}

filter_data_by_group <- function(data, group_name) {
    pattern <- paste (group_name, ".*", sep="")
    return(subset(data, select=grep(pattern, colnames(data))))
}

get_indices_by_names <- function(data, filter_colnames) {
    return(match(filter_colnames, colnames(data)))
}


# read data table from excel tab based on filename and sheet number
# calculate groups and add group labels to column headers
group_data <- function(data_file, sheet_num, grouping_function) {

    file_base <- sub("_out.xlsx", "", sub("out/", "", data_file))
    
    data <- get_sheet_first_data(data_file, sheet_num, "numeric")[,-1]
    data <- append_headers(data, rep(file_base, ncol(data)))
    data_groups <- apply(data, 2, grouping_function)
    grouped_data <- append_headers(data, data_groups)
    return(grouped_data)
}


get_dependent_data <- function(group_header, grouped_aktph, grouped_trpv){

    group_aktph_cells <- colnames(filter_data_by_group(grouped_aktph, group_header))
    group_aktph_indecis <- get_indices_by_names(grouped_aktph, group_aktph_cells)
    trpv_aktph <- grouped_trpv[, group_aktph_indecis]
    if (is.vector(trpv_aktph)) {
        trpv_aktph <- matrix(trpv_aktph)
        colnames(trpv_aktph) <- colnames(grouped_trpv)[group_aktph_indecis]
    }
    return(trpv_aktph)
}

build_grouped_data <- function(grouped_aktph, grouped_trpv){

    data_type_headers <- c("AKTPH", "TRPV", "TRPV-AKTPH")
    group_headers <- c("HIGH", "FLAT", "LOW")
    
    grouped_data_independent <- list(grouped_aktph, grouped_trpv)
    independent_headers <- data_type_headers[AKTPH_INDEX: TRPV_INDEX]
    names(grouped_data_independent) <- independent_headers

    grouped_data_dependent <- lapply(group_headers, 
        get_dependent_data, grouped_aktph=grouped_aktph, grouped_trpv=grouped_trpv)
    dependent_headers <- vapply(group_headers, 
        function(x) paste (data_type_headers[TRPV_AKTPH_INDEX], x, sep="."), "")
    names(grouped_data_dependent) <- dependent_headers
    
    grouped_data <- c(grouped_data_independent, grouped_data_dependent)
    return(grouped_data)
}


dump_into_file <- function(data, file) {
    output_file_name = paste(file, "_out.xlsx", sep="")
    
    logVar("    [START] Dumping results into", output_file_name)
    data_names <- names(data)
    options(warn=-1)
    file.remove(output_file_name)
    options(warn=0)
    lapply(data_names, function(x) 
        if (ncol(data[[x]]) > 0) {write.xlsx(data[[x]], file=output_file_name, sheetName=x, append=TRUE)})
    
    logVar("    [END] Dumping results into", output_file_name)
}
    
# Group cleansed data from step 1 related to baseline
grouping <- function(data_file, draw_plot_flag){

    out_base_name <- paste(sub(".xlsx", "", data_file), "_GRP", sep="")

    grouped_aktph <- group_data(data_file, AKTPH_INDEX, get_data_group_aktph)
    grouped_trpv <- group_data(data_file, TRPV_INDEX, get_data_group_trpv)
    
    grouped_data <- build_grouped_data(grouped_aktph, grouped_trpv)

    # optional, draw plot into R frame if input parameter draw_plot_flag == 1
    # TODO: unification, layout, PNG
    if (draw_plot_flag == 1 ) {
        plot_groups_independent(grouped_data, out_base_name)
    }
    
    if (DUMP_INTO_XLSX == 1 ) {
        dump_into_file(grouped_data, out_base_name)
    }

    return(grouped_data)
}


# Main processing function, encapsulating global parameters
# TODO: refactor for proper config usages.
trpv1_file <- function(experiment_file){

    data_file <- process_trpv1_input_file(experiment_file, DRAW_PLOTS)
    grouped_data <- grouping(data_file, DRAW_PLOTS)
    return(grouped_data)
}

if_vector_to_matrix <- function(v,old_m){
    if (is.vector(v)) {
        m <- matrix(v)
        colnames(m) <- colnames(old_m)
        return(m)
    } else {
        return(v)
    }
}

trpv1_files <- function(){
    
    experiment_files <- list.files(EXPERIMENT_FILES_DIR, pattern=".xlsx")
    grouped_data_list <- lapply(experiment_files, trpv1_file)
    summary <- grouped_data_list[[1]]
    summary_file_basename <- "out/trvp1_aggregated"

    for (grouped_data in grouped_data_list[-1]){
        for (name in names(summary) ){
            summ_data_matrix <- summary[[name]]
            grouped_data_matrix <- grouped_data[[name]]
            if (ncol(grouped_data_matrix) > 0)  {
                min_nrow <- min(nrow(summ_data_matrix), nrow(grouped_data_matrix))
                sum_part <- if_vector_to_matrix(summ_data_matrix[1:min_nrow,], summ_data_matrix)
                g_part <-  if_vector_to_matrix(grouped_data_matrix[1:min_nrow,], grouped_data_matrix)
                m <- cbind(sum_part, g_part)
                m <- m[, order(colnames(m))]
                summary[[name]] <- m
            }
        }
    }

    dump_into_file(summary, summary_file_basename)
    
    if (DRAW_PLOTS == 1 ) {
        plot_groups_independent(summary, summary_file_basename)
    }
}


################################# test run-trvp1.R #########################################
# trpv1_files()