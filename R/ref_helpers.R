.packageName <- "XLConnectXtra"

##' @title Convert reference in string or matrix form to vector form
##' @description Convert reference in string or matrix form to vector form.
##' @details Also check argument and results.
##' @param ref 
##' @return A numeric vector cell/area reference
##' @author Ben Veal
.ref2idx <- function(ref) {
    ## convert args to numeric vectors
    if(is.character(ref)) ref <- aref2idx(ref)
    if(is.matrix(ref) && nrow(ref) > 1) ref <- c(ref[1,],ref[2,])
    ## check args
    stopifnot(is.numeric(ref) && (length(ref) %in% c(2,4)))
    return(as.numeric(ref))
}

##' @title Check if a cell/region is within another region
##' @description Check if a cell/region is within another region.
##' @details This function returns TRUE if the first arg refers to a cell or region within the region referred to by the
##' second arg. Each arg can be either a vector of numeric row & column indices, or an excel reference/region string.
##' @param cell The index or aref of a single cell/region
##' @param region The index or aref of a region
##' @return TRUE if cell is within region
##' @author Joe Bloggs
##' @import XLConnect
##' @export 
isWithin <- function(cell,region) {
    ## check args and convert to numeric vectors    
    cell <- .ref2idx(cell)
    region <- .ref2idx(region)
    stopifnot(is.numeric(region) && length(region)==4)
    if((cell[1] >= region[1]) && (cell[1] <= region[3]) && (cell[2] >= region[2]) && (cell[2] <= region[4])) {
        if(length(cell)==2) return(TRUE)
        else return((cell[3] >= region[1]) && (cell[3] <= region[3]) && (cell[4] >= region[2]) && (cell[4] <= region[4]))
    } else return(FALSE)
}

##' @title Check if two regions overlap
##' @description Check if two regions overlap.
##' @details Each region argument can be either a numeric vector or matrix of row & column indices (as returned by aref2idx
##' or getReferenceCoordinatesForName) or an excel reference string (as returned by idx2aref).
##' Also the first argument (region1) may refer to a single cell instead of a region (i.e. it may be a vector of length 2).
##' @param region1 A vector of row & column indices or an excel reference string for a single cell or a region.
##' @param region2 A vector of row & column indices or an excel reference string for a region.
##' @return TRUE if region1 overlaps region2
##' @author Ben Veal
##' @export 
isOverlapping <- function(region1,region2) {
    ## check args and convert to numeric vectors
    region1 <- .ref2idx(region1)
    region2 <- .ref2idx(region2)
    ## region1 & region2 could be single points or pairs of points, need to check all cases:
    if(length(region1)==2) {
        if(length(region2)==2) all(region1==region2)
        else isWithin(region1,region2)
    } else if(length(region2)==2) {
        isWithin(region2,region1)
    } else { ## if both regions are pairs of points then only need to check 4 corners in total (2 from each)
        topright2 <- c(region2[1],region2[4])
        bottomleft2 <- c(region2[3],region2[2])
        isWithin(region1[1:2],region2) || isWithin(region1[3:4],region2) ||
            isWithin(topright2,region1) || isWithin(bottomleft2,region1)
    }
}

##' @title Return indices of cell/region adjacent to another region
##' @description Given a named region in an excel workbook this function can be used to calculate the cell indices
##' for a region located adjacent to the named one.
##' @details The return value will be either an excel reference string (as returned by XLConnect::aref), or a vector
##' of row & column indices (as returned by aref2idx, i.e. first 2 elements are row & column indices of top-left corner
##' and last 2 elements are row & column indices of bottom-right corner).
##' By default a reference to a single cell will be returned, but if size is set to something bigger than 1x1 then a
##' reference to a region will be returned. You can also add horizontal and vertical shifts to the region using the shift
##' argument. If values between 0 & 1 are used for the size argument then they will be multiplied by the height & width
##' of the named region to get the size values, similarly for the shift argument (which can take values between -1 & 1).
##' @param wb A workbook object
##' @param name The name of a named region in wb
##' @param location A string indicating which side of the named region to place the new region.
##' One of: "left","right","above"/"up","below"/"down" or just the first letter of one of those words.
##' @param shift A pair of integers indicating how much rightward & downward shift to add to the new region (default is c(0,0)).
##' Alternatively a pair of numbers between -1 & 1 indicating the amount to shift as a proportion of the named region size.
##' @param size A pair of integers or fractions indicating the horizontal & vertical length of the new region (default is c(1,1)).
##' Alternatively a pair of numbers between 0 & 1 indicating the size of the new region as a proportion of the named region size.
##' @param errorOnOverlap Logical indicating whether or not to throw an error if the new region overlaps the named region.
##' @param asAref Logical indicating whether to return the region as an excel reference string or a numeric vector.
##' @return A numeric vector or excel reference string of row & column indices for the new region.
##' @author Ben Veal
##' @export 
getRefsAdjacentToName <- function(wb,name,location="right",shift=c(0,0),size=c(1,1),errorOnOverlap=TRUE,asAref=FALSE) {
    stopifnot(class(wb)=="workbook",
              is.character(name) && length(name)==1,
              is.character(location) && length(location)==1,
              is.numeric(shift) && length(shift)==2,
              is.numeric(size) && length(size)==2,
              all(size >= 0))
    refs <- getReferenceCoordinatesForName(wb,name)
    ## convert fractional values of shift and size args to integer values
    for(i in 1:2) {
        for(arg in c("shift","size")) {
            vals <- get(arg)
            if(abs(vals[i]) > 0 && abs(vals[i]) < 1) {
                if(length(refs) > 2)
                    vals[i] <- round(vals[i]*(refs[2,i]-refs[1,i]))
                else
                    vals[i] <- 1
            }
            assign(arg,vals)
        }
    }
    ## calculate indices of new region depending on location
    if(grepl("^l",location,ignore.case=TRUE)) {
        rightidx <- refs[1,2] -1 + shift[2]
        leftidx <- rightidx - size[2] + 1
        topidx <- refs[1,1] + shift[1]
        bottomidx <- topidx + size[1] - 1
    } else if(grepl("^r",location,ignore.case=TRUE)) {
        leftidx <- refs[2,2] + 1 + shift[2]
        rightidx <- leftidx + size[2] - 1
        topidx <- refs[1,1] + shift[1]
        bottomidx <- topidx + size[1] - 1
    } else if(grepl("^(a|u)",location,ignore.case=TRUE)) {
        bottomidx <- refs[1,1] - 1 + shift[1]
        topidx <- bottomidx - size[1] + 1
        leftidx <- refs[1,2] + shift[2]
        rightidx <- leftidx + size[2] - 1
    } else if(grepl("^(b|d)",location,ignore.case=TRUE)) {
        topidx <- refs[2,1] + 1 + shift[1]
        bottomidx <- topidx + size[1] - 1        
        leftidx <- refs[1,2] + shift[2]
        rightidx <- leftidx + size[2] - 1
    } else stop("Invalid value for location arg:",location)
    if(topidx==bottomidx && leftidx==rightidx) idxs <- c(topidx,leftidx)
    else idxs <- c(topidx,leftidx,bottomidx,rightidx)
    if(errorOnOverlap && isOverlapping(idxs,refs))
        stop("New area overlaps adjacent area")
    if(leftidx < 1 || topidx < 1)
        stop("New area doesn't fit in worksheet")
    if(asAref) return(idx2aref(idxs)) else return(idxs)
}

##' @title Create function for writing to worksheet
##' @description Create a function for writing named regions to a predefined worksheet.
##' @details This function returns another convenience function for writing named regions to a given worksheet.
##' The returned function takes a region name, cell references and a dataframe as its main arguments and writes
##' the dataframe to the workbook & worksheet pair that were given as arguments to the parent function.
##' If the worksheet does not yet exist it is created by the parent function. Similarly if the named region does
##' not yet exist it is created by the child function.
##' The child function also has named arguments 'header' (see 'writeNamedRegion') and 'absRow' & 'absCol' (see 'idx2cref').
##' The default values for these arguments can be set using 'defaultHeader', 'defaultRow', and 'defaultCol' arguments of
##' the parent function.
##' @param wb An excel workbook object
##' @param sheetname The name of a worksheet in the workbook object wb
##' @param defaultRow Default logical value to use for absRow arg of returned function which indicates whether
##' rows of cellrefs should be absolute references or not.
##' @param defaultCol Default logical value to use for absCol arg of returned function which indicates whether
##' columns of cellrefs should be absolute references or not.
##' @param defaultHeader Default logical value of header arg of returned function which indicates whether column
##' names of data should be written to the excel file.
##' @author Ben Veal
##' @export 
makeRegionFunction <- function(wb,sheetname,defaultRow=TRUE,defaultCol=TRUE,defaultHeader=TRUE) {
    ## check args
    stopifnot(class(wb)=="workbook",
              is.character(sheetname) && length(sheetname)==1,
              is.logical(defaultRow),
              is.logical(defaultCol),
              is.logical(defaultHeader))
    ## create sheet if it doesn't already exist
    if(!(existsName(wb,sheetname))) createSheet(wb,name=sheetname)
    ## return function for writing named regions
    function(name,cellrefs,data,header=defaultHeader,absRow=defaultRow,absCol=defaultCol) {
        removeName(wb,name)
        if(is.numeric(cellrefs) && ((length(cellrefs)==2)||length(cellrefs)==4)) {
            cellrefs <- paste(idx2cref(cellrefs,absRow=absRow,absCol=absCol),collapse=":")
        } else if(!is.character(cellrefs) || length(cellrefs)!=1) {
            stop("Invalid cellrefs argument: ",cellrefs)
        } 
        createName(wb,name=name,formula=paste0(sheetname,"!",cellrefs))
        writeNamedRegion(wb,data=data,name=name,header=header)
    }
}

##' @title Create an absolute reference from base and relative references
##' @description Given a cell reference that is relative to a given base reference, create a corresponding absolute cell reference.
##' @details The reference arguments can refer to single cells or areas. 
##' The size of the returned reference area will correspond with the size of the relref argument, and if the relref argument
##' is a single cell then the returned reference will also be a single cell.
##' Reference arguments can be either numeric vectors or matrices of row & column indices (as returned by aref2idx
##' or getReferenceCoordinatesForName) or a excel reference strings (as returned by idx2aref).
##' A relative reference of 1 corresponds to the first row/column.
##' Negative indices in the relfref argument count backwards from the bottom/right side of the base reference cell/area,
##' with 0 corresponding to the last row/column.
##' @param baseref An absolute reference to an excel worksheet cell/area in either string form, matrix form or numeric vector form.
##' @param relref A relative reference to an excel worksheet cell/area in either string form, matrix form or numeric vector form.
##' @return A vector of 2 or 4 row & column indices for a single cell or an area, with entries 1 & 3 indicating the first & last
##' rows, and entries 2 & 4 indicating the first and last columns.
##' @author Ben Veal
##' @examples rel2absref("B2:E5","A1:B2")
##' rel2absref("B2:E5",c(1,1,2,2))
##' rel2absref("B2:E5",c(1,1,-2,-2))
##' rel2absref("B2:E5",c(-3,-3,-2,-2))
##' @export 
rel2absref <- function(baseref,relref) {
    ## check args and convert to numeric vectors
    baseref <- .ref2idx(baseref)
    relref <- .ref2idx(relref)
    absref <- integer(length(relref))
    ## first row and column
    for(i in c(1,2)) {
        if(relref[i] > 0) absref[i] <- baseref[i] + relref[i] - 1
        else if(length(baseref) > 2) absref[i] <- baseref[i+2] + relref[i]
        else stop("invalid args")
    }
    ## last row and column
    if(length(relref) > 2) {
        for(i in c(3,4)) {
            if(relref[i] > 0) absref[i] <- baseref[i-2] + relref[i] - 1
            else if(length(baseref) > 2) absref[i] <- baseref[i] + relref[i]
            else stop("invalid args")
        }
    }
    ## check that final indices make sense and return them
    if((length(absref) > 2 && (absref[1] > absref[3] || absref[2] > absref[4])) || any(absref < 1))
        stop("Invalid area!")
    return(absref)
}
