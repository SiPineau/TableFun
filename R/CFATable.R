
#' CFATable
#'
#'@author Simon Pineau
#'
#'@description
#'Create an APA table from lavaan cfa fit object.
#'
#' @param CFAfit A lavaan cfa fit object.
#' @param header Name of the table in quotes.
#' @param digits Number of decimal places in the table. By default digits = 3.
#' @param footer Optional argument. Inserts a general table footnote. The argument must be in quotes.
#' @param adjtopage Set columns width to default page margins (2.5) and set space between text and lines.This argument must be a vector composed of: orientation of page ("landscape/portrait"), does first column need to be wider (True/False), if previews element set to TRUE set ratio of first column size (numeric greater than 0), space between lines and text (numeric, in points). Default is c("portrait", F,1,3).
#' @param CorMatrix Include the factor correlation matrix in the table. Default = FALSE.
#' @param FCorMat Include a specific factor correlation matrix (for exemple after factor score computation). If FCorMat=NULL and CorMatrix=TRUE, psi matrix from CFAfit.
#' @param savename Optional argument. Saves the table in .docx format in the working space. The argument must be in quotes.
#'
#' @import rempsyc flextable export officer stringr dplyr tidyr rlang
#' @return CFA APA Table
#' @export
#' @examples
#'
#' cfamodel <- '
#'   f1 ~= V1 + V2 + V3
#'   f2 ~= V4 + V5 + V6
#' '
#'
#' cfafit <- cfa(model = cfamodel, data = df)
#'
#' CFATable(cfafit, "CFA APA table")

CFATable <- function(CFAfit, header, digits = 3, footer, adjtopage = c("portrait", F,1,3), CorMatrix = F, FCorMat = NULL, savename){

  cfatable <- data.frame(lavInspect(CFAfit, what = "std")$lambda)

  cfatable[cfatable == 0] <- NA

  cfatable <- cfatable %>% mutate_if(is.numeric, round, digits)

  if(CorMatrix == TRUE | !missing(FCorMat)){
    cfatable <- cfatable %>% mutate(Items = rownames(.), Section = "Factor loadings") %>% relocate(Section, Items)
    if(!missing(FCorMat)){
      FCorMat[upper.tri(FCorMat, diag = F)] <- NA

      cfapsitable <- data.frame(CorMatrix)

    }else{
      cfapsitable <- lavInspect(CFAfit, what = "std")$psi

      cfapsitable[upper.tri(cfapsitable, diag = F)] <- NA

      cfapsitable <- data.frame(cfapsitable)

    }
    cfapsitable <- cfapsitable %>% mutate_if(is.numeric, round, digits)

    cfapsitable <- rbind(colnames(cfapsitable), cfapsitable)

    cfapsitable <- cfapsitable %>% mutate(Items = rownames(.), Section = "Correlation matrix") %>% relocate(Section, Items)

    cfapsitable[1,2] = " "

    cfatable <- rbind(cfatable, cfapsitable)

    cfatable$Section <- factor(cfatable$Section, levels = c("Factor loadings", "Correlation matrix"))

    TF <- TableFun(cfatable, header = header, adjtopage = adjtopage, spread_col = c("Section", "Items", "center"), digits = digits, savename = savename, left.align = 1, footer = footer)

  }else{
    cfatable <- cfatable %>% mutate(Items = rownames(.)) %>% relocate(Items)

    colnames(cfatable) <- paste("Factor loadings", colnames(cfatable), sep = "_")

    TF <- TableFun(cfatable, header = header, adjtopage = adjtopage, digits = digits, savename = savename, left.align = 1, footer = footer)
  }
  return(TF)
}
