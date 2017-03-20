.packageName <- "XLConnectXtra"


##' @title Create formulas for linear combinations
##' @description Create formulas for linear combinations of values in columns.
##' @details A formula will be created for each row of the area indicated by the vars argument.
##' Each formula has terms which are products the columns of the area indicated by the vars argument with either
##' the corresponding coefficient in the coeffvals argument, or if that is not present then the value in the corresponding cell
##' indicated by the coeffrefs argument. These terms are summed together to produce the final formula.
##' @param vars A vector or matrix of row & column indices or an excel reference string for the region whose columns contain
##' the variables corresponding with the coefficients 
##' @param coeffvals a numeric vector of coefficient values
##' @param coeffrefs a vector or matrix or row & column indices or an excel reference string for the area of the worksheet
##' containing the coefficients
##' @return A character vector containing the linear equations
##' @author Ben Veal
##' @export 
createLinCombFormula <- function(vars,coeffvals=NULL,coeffrefs=NULL) {
    ## coeffs can be a string containing an aref so convert to indices
    if(!is.null(coeffvals)) {
        stopifnot(is.numeric(coeffvals))
        coeffs <- coeffvals
    } else if(!is.null(coeffrefs)) {
        coeffs <- .ref2idx(coeffrefs)
        if(coeffs[1]==coeffs[3])
            coeffs <- paste0(idx2col(coeffs[2]:coeffs[4]),coeffs[1])
        else if(coeffs[2]==coeffs[4])
            coeffs <- paste0(idx2col(coeffs[2]),coeffs[1]:coeffs[3])
        else stop("Invalid coeffrefs argument!")
    } else stop("Missing coeffvals or coeffrefs argument!")
    ## check vars and convert to a numeric vector
    vars <- .ref2idx(vars)
    varrows <- vars[1]:vars[3]
    varcols <- sapply(vars[2]:vars[4],idx2col)
    ## create the first term 
    linforms <- paste0("(",varcols[1],varrows,"*",coeffs[1],")")
    ## create subsequent terms
    for(i in 2:length(coeffs))
        linforms <- paste(linforms,paste0("(",varcols[i],varrows,"*",coeffs[i],")"),sep="+")
    return(linforms)
}

## TODO: Make generic function for creating excel formulas from model objects (e.g. lm, glm, polr, rpart)
## Use S3 generic methods, see my R notes and maybe use library(GString) and library(R.oo). 
## createExcelFormulas.lm <- function(lm) {
## }
