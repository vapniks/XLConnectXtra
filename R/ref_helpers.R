.packageName <- "XLConnectXtra"


##' @title Check if a cell/region is within another region
##' @description Check if a cell/region is within another region
##' @details This function returns TRUE if the first arg refers to a cell or region within the region referred to by the
##' second arg. Each arg can be either a vector of numeric row & column indices, or an excel reference/region string.
##' @param cell The index or aref of a single cell/region
##' @param region The index or aref of a region
##' @return TRUE if cell is within region
##' @author Joe Bloggs
##' @import XLConnect
##' @export 
isWithin <- function(cell,region) {
    if(is.character(cell)) cell <- aref2idx(cell)    
    if(is.character(region)) region <- aref2idx(region)
    if(is.matrix(cell) && nrow(cell) > 1) cell <- c(cell[1,],cell[2,])
    if(is.matrix(region) && nrow(region) > 1) region <- c(region[1,],region[2,])
    stopifnot(is.numeric(cell) && length(cell) %in% c(2,4),
              is.numeric(region) && length(region)==4)
    if((cell[1] >= region[1]) && (cell[1] <= region[3]) && (cell[2] >= region[2]) && (cell[2] <= region[4])) {
        if(length(cell)==2) return(TRUE)
        else return((cell[3] >= region[1]) && (cell[3] <= region[3]) && (cell[4] >= region[2]) && (cell[4] <= region[4]))
    } else return(FALSE)
}

#### function that checks if two regions overlap, returns TRUE if they do
isOverlapping <- function(region1,region2) {
    ## convert args to numeric vectors
    if(is.character(region1)) region1 <- aref2idx(region1)
    if(is.character(region2)) region2 <- aref2idx(region2)
    if(is.matrix(region1) && nrow(region1) > 1) region1 <- c(region1[1,],region1[2,])
    if(is.matrix(region2) && nrow(region2) > 1) region2 <- c(region2[1,],region2[2,])
    ## check args
    stopifnot(is.numeric(region1) && (length(region1) %in% c(2,4)),
              is.numeric(region2) && (length(region2) %in% c(2,4)))
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
#### function that returns the idx of cell adjacent to named region 
getRefsAdjacentToName <- function(wb,name,location="right",shift=c(0,0),size=c(1,1),errorOnOverlap=TRUE,asAref=FALSE) {
    stopifnot(class(wb)=="workbook",
              is.character(name) && length(name)==1,
              is.character(location) && length(location)==1,
              is.numeric(shift) && length(shift)==2,
              is.numeric(size) && length(size)==2,
              size[1] > 0, size[2] > 0)
    refs <- getReferenceCoordinatesForName(wb,name)
    if(location=="left") {
        rightidx <- refs[1,2] -1 + shift[2]
        leftidx <- rightidx - size[2] + 1
        topidx <- refs[1,1] + shift[1]
        bottomidx <- topidx + size[1] - 1
    } else if(location=="right") {
        leftidx <- refs[2,2] + 1 + shift[2]
        rightidx <- leftidx + size[2] - 1
        topidx <- refs[1,1] + shift[1]
        bottomidx <- topidx + size[1] - 1
    } else if(location=="above") {
        bottomidx <- refs[1,1] - 1 + shift[1]
        topidx <- bottomidx - size[1] + 1
        leftidx <- refs[1,2] + shift[2]
        rightidx <- leftidx + size[2] - 1
    } else if(location=="below") {
        topidx <- refs[2,1] + 1 + shift[1]
        bottomidx <- topidx + size[1] - 1        
        leftidx <- refs[1,2] + shift[2]
        rightidx <- leftidx + size[2] - 1
    } else stop("Invalid value for location arg:",location)
    if(topidx==bottomidx && leftidx==rightidx) idxs <- c(topidx,leftidx)
    else idxs <- c(topidx,leftidx,bottomidx,rightidx)
    if(isOverlapping(idxs,refs))
        stop("New area overlaps adjacent area")
    if(leftidx < 1 || topidx < 1)
        stop("New area doesn't fit in worksheet")
    if(asAref) return(idx2aref(idxs)) else return(idxs)
}

### write data to worksheet
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
