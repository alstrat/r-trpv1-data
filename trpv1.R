rm(list=ls(all=TRUE))

library("xlsx")

logVar <- function(string, var) {
    print(paste(string, " ",var))
}

reduce_bg_noise <- function(cell) {
    treshhold = 0
    result = cell

    if (cell > treshhold ) { 
        result = cell
    } 
    return(result)
}

process_trpv_col <- function(in_trpv_column, bg, time, noise_index) {

    trpv_column_no_bg_reduced <- vapply((in_trpv_column - bg), reduce_bg_noise, 0)
    trpv_column <- trpv_column_no_bg_reduced

    y_lm <- trpv_column[1:noise_index] 
    x_lm <- time[1:noise_index]
    lin_reg <- lm(formula = y_lm ~ x_lm)
    
    intercept_lm <- summary(lin_reg)$coefficients[1,1]

    x_lm_predict <- time[1:(length(time))]
    new_x_lm <- data.frame("x_lm"=x_lm_predict)
    new_y_lm <- predict(lin_reg, newdata=new_x_lm)
    
    cleansed_trpv_column <- trpv_column - new_y_lm + rep(intercept_lm, (length(time)))
    
    mean_noaction <- mean(cleansed_trpv_column[1:20])
    
    normalized_trpv <- vapply(cleansed_trpv_column, function(x) x/mean_noaction, 0)
    
    return(normalized_trpv)
}


get_sheet_data <- function(file_name, sheet_number, bg_index, noise_index){
    raw <- read.xlsx2(file_name, 
            sheetIndex = sheet_number, 
            colIndex = (1:bg_index),
                  colClasses=rep("numeric",bg_index))

    bg <- data.matrix(raw[bg_index])[,1]
    time <- data.matrix(raw[1])[,1]
    trpv_data <- data.matrix(raw[2:(bg_index-1)])
    
    processed_data <- apply(trpv_data, 2, process_trpv_col, bg=bg, time=time, noise_index=noise_index)

    return(processed_data)
}


is_significant_column <- function(data_column, split) {
    treshhold = 1.2
    result = 0

    action = data_column[split: (2*split)]
    action_criteria = max(action)
    
    if (action_criteria > treshhold) { 
        result = 1
    } 
    return(result)
}

filter_data <- function(trpv1, aktph, noise_index, flag) {
    significant_column_statuses = apply(aktph, 2, is_significant_column, split=noise_index)
    significant_columns <- significant_column_statuses[significant_column_statuses != 0]
    
    filtered = subset(aktph, select=names(significant_columns))

    if (flag == 1) { 
        filtered = subset(trpv1, select=names(significant_columns))
    }
    return(filtered)
}


plot_trvp <- function(file_name, trpv1, aktph, filtered_trvp, filtered_aktph) {
# row 1 is 1/3 the height of row 2
# column 2 is 1/4 the width of the column 1
# attach(mtcars)
# layout(matrix(c(1,1,2,3), 2, 2, byrow = TRUE),
   # widths=c(3,1), heights=c(1,2))
   
   png_output_file_name = paste(file_name, "_out.png", sep="")
   
   logVar("    [START] Building plots into ", png_output_file_name)
   
    png(png_output_file_name,
        width     = 18,
        height    = 9,
        units     = "in",
        res       = 1200,
        pointsize = 4
    )
     # png(png_output_file_name,
        # width     = 9,
        # height    = 4.5,
        # units     = "in",
        # res       = 120,
        # pointsize = 4
    # )
    par(c(2,2),
        xaxs     = "i",
        yaxs     = "i",
        cex.axis = 2,
        cex.lab  = 2
    )
    layout(matrix(c(1,2,3,4), 2, 2, byrow = TRUE),
            widths=c(1,1), heights=c(1,2))
    
    matplot(trpv1, type = "l")
    m_names = names(trpv1[1,])
    nn <- ncol(trpv1)
    legend("bottomleft", legend = m_names,col=seq_len(nn),cex=2, fill=seq_len(nn))
    
    matplot(aktph, type = "l")
    m_names = names(aktph[1,])
    nn <- ncol(aktph)
    legend("bottomleft", legend = m_names,col=seq_len(nn),cex=2, fill=seq_len(nn))
    
    matplot(filtered_trvp, type = "l")
    m_names = names(filtered_trvp[1,])
    nn <- ncol(filtered_trvp)
    legend("bottomleft", legend = m_names,col=seq_len(nn),cex=2, fill=seq_len(nn))
    
    matplot(filtered_aktph, type = "l")
    m_names = names(filtered_aktph[1,])
    nn <- ncol(filtered_aktph)
    legend("bottomleft", legend = m_names,col=seq_len(nn),cex=2, fill=seq_len(nn))
    dev.off()
    
    logVar("    [END] Building plots into ", png_output_file_name)
 }


trpv_file <- function(in_file_name, in_bg_index, in_noise_index) {

    file_name <- paste("in/", in_file_name, sep="")
    out_base_name <- paste("out/", (sub(".xlsx", "", in_file_name)), sep="")
    
    noise_index <- as.numeric(in_noise_index)
    bg_index <- as.numeric(in_bg_index)

    logVar("[START] Processing", file_name)
    logVar("    BG index", bg_index)
    logVar("    index where noise ends", noise_index)
    
    trpv1 <- get_sheet_data(file_name, 2, bg_index, in_noise_index)
    print("trv done")
    aktph <- get_sheet_data(file_name, 1, bg_index, in_noise_index)

    filtered_trvp <- filter_data(trpv1, aktph, noise_index, 1)
    filtered_aktph <- filter_data(trpv1, aktph, noise_index, 2)


    plot_trvp(out_base_name, trpv1, aktph, filtered_trvp, filtered_aktph)
   
    output_file_name = paste(out_base_name, "_out.xlsx", sep="")
    logVar("    [START] Dumping results into", output_file_name)
    
    
    write.xlsx(trpv1, file=output_file_name, sheetName="processed_trpv1")
    write.xlsx(filtered_trvp, file=output_file_name, sheetName="reduced_trpv1", append=TRUE)
    write.xlsx(aktph, file=output_file_name, sheetName="processed_aktph", append=TRUE)
    write.xlsx(filtered_aktph, file=output_file_name, sheetName="reduced_aktph", append=TRUE)
    logVar("    [END] Dumping results into", output_file_name)
    
    logVar("[END] Processing", file_name)
    
}


################################# run-trvp1.R #########################################

global_noise_index = 120

# Process files
trpv_file("04-19-16_slip1.xlsx",11, global_noise_index)
trpv_file("04-13-16_slip1.xlsx",15, global_noise_index)
