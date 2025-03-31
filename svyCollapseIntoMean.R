svyCollapseIntoMean <- function(x, do_it = F){
  if(!do_it){return(x)}
  
  out <- x
  out[1, ] <- 0
  out[2, ] <- colSums(x * (rownames(x) %>% as.numeric)) / colSums(x)
  out <- out[1:2, ]
  if(!is.null(rownames(out))){
    rownames(out) <- c("FALSE", "TRUE")  
  } else {
    names(out) <- c("FALSE", "TRUE")
  }
  
  return(out %>% round(2))
}
