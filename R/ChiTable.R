
#' ChiTable
#'
#'@author Simon Pineau
#'
#'@description
#'Create an APA norme chi² table.
#'
#' @param chitest A chi² object from chisq.test function
#' @param header Name of the table in quotes.
#' @param digits Number of decimal places in the table. By default digits = 2.
#' @param expected Does the table need to include expected repartition. Default = FALSE.
#' @param footer Optional argument. Inserts a general table footnote. The argument must be in quotes.
#' @param adjtopage Set columns width to default page margins (2.5) and set space between text and lines.This argument must be a vector composed of: orientation of page ("landscape/portrait"), does first column need to be wider (True/False), if previews element set to TRUE set ratio of first column size (numeric greater than 0), space between lines and text (numeric, in points). Default = adjtopage = c("portrait", F, 1, 3).
#' @param savename Optional argument. Saves the table in .docx format in the working space. The argument must be in quotes.
#'
#' @import rempsyc flextable export officer stringr dplyr tidyr rlang
#' @return APA Table
#' @export
#' @examples
#'
#' chi2<-chisq.test(table(df[c("V1","V2")]))
#'
#' ChiTable(chi2, header = "chi2 table", expected = F, adjtopage = c("portrait", F, 1, 3))

ChiTable <- function(chitest, header, digits = 2, expected = FALSE, footer, adjtopage, savename){
  if(expected == FALSE){
    chiobs <- data.frame(chitest$observed)
    chistd <- data.frame(chitest$stdres)

    chistd <-  chistd %>% rename("Res" = "Freq") %>%
      mutate(across(Res, \(x) round(x, 2)))

    chifull <- full_join(chiobs, chistd, by = c(colnames(chiobs)[1], colnames(chiobs)[2]))

    chifull <- chifull %>%
      mutate(value=paste(Freq, " (", Res, ")", sep = "")) %>%
      select(-Res, - Freq) %>%
      pivot_wider(names_from = colnames(chiobs)[1], values_from=value, names_prefix = paste(colnames(chiobs)[1], "_", sep = "")) %>%
      arrange(colnames(chiobs)[2])

    TableFun(chifull, header = header, digits = digits, merge.col = 1, left.align = 1, adjtopage = adjtopage, footer = footer, savename = savename)

  }else{

    chiobs <- data.frame(chitest$observed)
    chiexp <- data.frame(chitest$expected)
    chistd <- data.frame(chitest$stdres)


    colnames(chiexp) <- as.numeric(unique(chiobs[,2]))
    chiexp <- chiexp %>% mutate(!!colnames(chiobs)[1] := row.names(chiexp)) %>% gather(., key = !!colnames(chiobs)[2], value = "Expected", -!!colnames(chiobs)[1])

    chiobs <-  chiobs %>% rename("Observed" = "Freq")

    chistd <-  chistd %>% rename("Residuals" = "Freq")

    chifull <- Reduce(function(x, y) full_join(x, y, by = c(colnames(chiobs)[1], colnames(chiobs)[2])), list(chiobs, chiexp, chistd))


    chifull <- chifull %>% gather(., key = "Ind", value = "value", -c(colnames(chiobs)[1], colnames(chiobs)[2]))

    chifull <- chifull %>%
      pivot_wider(names_from = colnames(chiobs)[1], values_from=value, names_prefix = paste(colnames(chiobs)[1], "_", sep = ""))

    chifull[,colnames(chiobs)[2]]<-factor(chifull[[colnames(chiobs)[2]]])

    chifull <- chifull %>% arrange(.data[[colnames(chiobs)[2]]])

    TableFun(chifull, header = header, digits = digits, spread_col = c(colnames(chifull)[1], colnames(chifull)[2], "left"), adjtopage = adjtopage, footer = footer, savename = savename)

  }
}
