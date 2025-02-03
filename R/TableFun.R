
#' TableFun
#'
#'@author Simon Pineau
#'
#'@description
#'Create an APA norme table.
#'
#' @param df A dataframe.
#' @param header Name of the table in quotes.
#' @param digits Number of decimal places in the table. By default digits = 2.
#' @param left.align Optional argument. Aligns the information in the body of the table to the left.
#' @param right.align Optional argument. Aligns the information in the body of the table to the right.
#' @param top.align Optional argument. Aligns the information in the body of the table at the top.
#' @param bottom.align Optional argument. Aligns the information in the body of the table at the bottom.
#' @param merge.col Optional argument. Merges adjacent similar rows of specified columns.
#' @param footer Optional argument. Inserts a general table footnote. The argument must be in quotes.
#' @param foot.note.header Optional argument. Inserts a specific note attached to a column title. This argument must be a vector composed of: the column number, the "letter" associated with the note, and the "specific note": c(1, "a", "specific note"). Several specific notes can be added as follow: c(1, "a", "specific note1", 2, "b", "specific note2").
#' @param foot.note.body Optional argument. Inserts a specific note attached to a cell in the body of the table. The argument is written the same way as foot.note.header by adding the row number after the column number: c(1, 1, "a", "specific note")
#' @param H.rotate Optional argument. Select columns to rotate the header's text.
#' @param B.rotate Optional argument. Select columns to rotate the body's text.
#' @param col.bg.color/row.bg.color Optional argument. Add a color to the background of selected columns/row. This argument must be a vector composed of: column's/row's number, "color code", part of the table affected ("header" / "body" / "all").
#' @param spread_col Optional argument. Spread one column values as row separator based on an other character column. This argument must be a vector composed of: "column name to be spread", "column name to be separate", postition of spreaded column value : "center" / "left"
#' @param savename Optional argument. Saves the table in .docx format in the working space. The argument must be in quotes.
#'
#' @import rempsyc flextable export officer stringr dplyr tidyr rlang
#' @return
#' @export
#' @examples
#' # Simple table
#' df<-data.frame("Variables" = c(1,2,3),
#'                "Mean" = c(32,22,15),
#'                "SD" = c(0.2,0.8,0.4))
#'
#' TableFun(df, header = "Mean and sd", foot.note.header = c(3,"a","SD = Standard deviation"))
#'
#' # Multiple colmuns names level
#'
#' df<-data.frame("Variables" = c(1,2,3),
#'                "T1_Mean" = c(32,22,15),
#'                "T1_SD" = c(0.2,0.8,0.4),
#'                "T2_Mean" = c(30,24,18),
#'                "T2_SD" = c(0.3,0.5,0.6))
#'
#' TableFun(df,header = "Mean and sd across 2 times points", foot.note.header = c(3,"a","ET = Standard deviation"))
#'
#' # Spreaded column
#'
#' df<-data.frame("Variables" = c(1,2,3),
#'                "T1_Mean" = c(32,22,15),
#'                "T1_SD" = c(0.2,0.8,0.4),
#'                "T2_Mean" = c(30,24,18),
#'                "T2_SD" = c(0.3,0.5,0.6))
#'
#' df <- reshape(df, varying = c(2:5), idvar = "Variables", sep = "_", timevar = "ind", direction = "long") %>% arrange(Variables)
#'
#' TableFun(df,header = "Mean and sd across 2 times points", spread_col = c("Variables", "ind", "left"))


TableFun <- function(df, header, digits = 2, left.align = NULL, right.align = NULL, top.align = NULL, bottom.align = NULL, merge.col = NULL, footer = NULL, foot.note.header = NULL, foot.note.body = NULL, H.rotate = NULL, B.rotate = NULL, col.bg.color = NULL, row.bg.color = NULL, spread_col = NULL, savename = NULL){

  df <- df %>% data.frame(., check.names = F) %>%
    mutate_if(is.numeric, round, digits = digits)

  if(!missing(spread_col)){
    if(length(spread_col)<3){
      stop("Missing argument in 'spread_col'")
    }
      namevar <- spread_col[c(1,2)]

      if(length(colnames(df[,c(namevar)])[sapply(df[,c(namevar)], is.factor)]) > 0){
        fctvar <- colnames(df[,c(namevar)])[sapply(df[,c(namevar)], is.factor)]
        fctlvl <- sapply(df[,c(fctvar)], levels)

        for (f in 1:length(fctvar)) {
          df[,fctvar[f]] <- as.character(df[,fctvar[f]])
        }
      }

      indexvar <- match(namevar, names(df))

      nbv<-data.frame(table(df[,indexvar[1]]))

      nbcol <- ncol(df)

      for (v in 1:nrow(nbv)) {
        if(v == 1){
          tmp <- rbind(rep(as.character(nbv[v,1]), nbcol), df %>% filter(!!sym(namevar[1]) == nbv[v,1]))
        }else{
          tmp[c((nrow(tmp)+1):((nrow(tmp)+1)+nbv[v,2])),] <- rbind(rep(as.character(nbv[v,1]), nbcol), df %>% filter(!!sym(namevar[1]) == nbv[v,1]))
        }
      }

      Newnbv <- data.frame(table(tmp[,indexvar[1]]))

      for (v in 1:nrow(Newnbv)) {
        if(v == 1){
          newrow <- c(1)
        }else{
          newrow <- c(newrow,(sum(Newnbv[c(1:(v-1)),2])+1))
        }
      }

      if(spread_col[3] == "center"){
        firstj <- 1
      }

      if(spread_col[3] == "left"){
        firstj <- 2
      }

      if(!missing(fctlvl)){
        tmp[,namevar[2]] <- paste(tmp[,namevar[1]], tmp[,namevar[2]], sep="_")
        for (lvl1 in 1:length(fctlvl[[1]])) {
          for (lvl2 in 1:length(fctlvl[[2]])) {
            if(lvl1*lvl2 == 1){
              newfctlvl <- paste(fctlvl[[1]][lvl1], fctlvl[[2]][lvl2], sep = "_")
            }else{
              newfctlvl[(length(newfctlvl)+1)] <- paste(fctlvl[[1]][lvl1], fctlvl[[2]][lvl2], sep = "_")
            }
          }
        }

        tmp[,namevar[2]] <- factor(tmp[,namevar[2]], levels = c(newfctlvl))

        tmp <- tmp %>% data.frame(., check.names = F) %>% mutate(!!sym(namevar[2]) := case_when(!!sym(namevar[2]) := is.na(!!sym(namevar[2])) ~ paste(!!sym(namevar[1]), !!sym(namevar[1]), sep = "_"),
                                                                                                .default = !!sym(namevar[2])))

        tmp[,namevar[1]] <- factor(tmp[,namevar[1]], levels = c(fctlvl[[1]]))

        tmp <- tmp %>% arrange(!!sym(namevar[1]))

        tmp <- tmp %>% separate(!!sym(namevar[2]), c("Vartodelete", namevar[2]), sep = "_") %>%
          select(-Vartodelete)
      }

      df1 <- tmp  %>%
        select(-!!sym(namevar[1])) %>%
        rename(!!sym(namevar[1]) := !!sym(namevar[2]))

  }
  x<-flextable(df) %>%
    separate_header() %>%
    autofit()

  if(!missing(merge.col)){
    x<-merge_v(x, j = merge.col, combine = TRUE)
  }

  if(!missing(footer) | !missing(foot.note.header) | !missing(foot.note.body)){
    x<-add_footer_lines(x,values = "")

    if(missing(footer)){
      x<-compose(x,i = 1, j = 1, value = as_paragraph(as_i("Note. ")), part = "footer")

    } else {

      x<-compose(x,i = 1, j = 1, value = as_paragraph(as_i("Note. "), footer), part = "footer")
    }
  }

  if(!missing(foot.note.header)){
    y = 1
    x<-footnote(x, j = as.double(foot.note.header[y]), part = "header", ref_symbols = foot.note.header[y+1], value = as_paragraph(str_c(foot.note.header[y+2], ". ")))

    while(y < length(foot.note.header)-2){
      y = y+3
      x<-footnote(x, j = as.double(foot.note.header[y]), part = "header", ref_symbols = foot.note.header[y+1], value = as_paragraph(foot.note.header[y+2]), inline = TRUE, sep = ". ")
    }
  }

  if(!missing(foot.note.body)){
    z = 1
    x<-footnote(x, i = as.double(foot.note.body[z]), j = as.double(foot.note.body[z+1]), part = "body", ref_symbols = foot.note.body[z+2], value = as_paragraph(foot.note.body[z+3]), inline = TRUE, sep = ". ")

    while(z < length(foot.note.body)-3){
      z = z+4
      x<-footnote(x, i = as.double(foot.note.body[z]), j = as.double(foot.note.body[z+1]), part = "body", ref_symbols = foot.note.body[z+2], value = as_paragraph(foot.note.body[z+3]), inline = TRUE, sep = ". ")
    }
  }

  x <- x %>%
    fontsize(part = "all", size = 11) %>%
    add_header_lines(header) %>%
    italic(i = 1, part = "header") %>%
    hline_top(part = "header",
              border = fp_border(color = "black",
                                 width = 0,
                                 style = "solid")) %>%
    hline(i = c(1,2),
          part = "header",
          border = fp_border(color = "black",
                             width = 0.25,
                             style = "solid")) %>%
    hline_top(part = "body",
              border = fp_border(color = "black",
                                 width = 0.25,
                                 style = "solid")) %>%
    hline_bottom(part = "body",
                 border = fp_border(color = "black",
                                    width = 0.25,
                                    style = "solid")) %>%

    align(part = "all", align = "center") %>%
    valign(j = setdiff(c(1:length(x$col_keys)), which(grepl("[_\\.]", x$col_keys))), valign = "top", part = "header") %>%
    align(i = 1, part = "header", align = "left") %>%
    align(i = 1, part = "footer", align = "left")

  if(!missing(top.align)){
    x<-valign(x,j = c(top.align), part = "body", valign = "top")
  }

  if(!missing(bottom.align)){
    x<-valign(x,j = c(bottom.align), part = "body", valign = "bottom")
  }

  if(!missing(left.align)){
    x<-align(x,j = c(left.align), part = "body", align = "left")
  }

  if(!missing(right.align)){
    x<-align(x,j = c(right.align), part = "body", align = "right")
  }

  if(!missing(H.rotate)){
    x <- rotate(x, j = c(as.integer(H.rotate[-length(H.rotate)])), rotation = H.rotate[length(H.rotate)], part = "header")
  }

  if(!missing(B.rotate)){
    x <- rotate(x, j = c(as.integer(B.rotate[-length(B.rotate)])), rotation = B.rotate[length(B.rotate)], part = "body")
  }

  if(!missing(col.bg.color)){
    x<-bg(x, j = as.integer(col.bg.color[1:(length(col.bg.color)-2)]), bg = col.bg.color[length(col.bg.color)-1], part = col.bg.color[length(col.bg.color)])
  }

  if(!missing(row.bg.color)){
    x<-bg(x, i = c(as.integer(row.bg.color[1:(length(row.bg.color)-2)])), bg = row.bg.color[length(row.bg.color)-1], part = row.bg.color[length(row.bg.color)])
  }

  if(!missing(spread_col)){

    x <- x %>%
      merge_h(., i = c(newrow), part = "body") %>%
      align(part = "body", align = "left") %>%
      align(i = c(1:nrow(tmp)), j = c(firstj:(nbcol-1)), part = "body", align = "center") %>%
      align(i = setdiff(c(1:nrow(tmp)), newrow), j = 1, part = "body", align = "left") %>%
      prepend_chunks(i = setdiff(c(1:nrow(tmp)), newrow), j = namevar[1], as_chunk("\t"), part = "body")

    if(spread_col[3] == "center"){
      newrowborder <- newrow
      for (v in 2:length(newrow)) {
        newrowborder[(length(newrowborder)+1)] <- (newrow[v] - 1)
      }
      x <- x %>%
        hline(i = newrowborder,
              part = "body",
              border = fp_border(color = "black",
                                 width = 0.25,
                                 style = "solid"))
    }
  }

x <- fix_border_issues(x,part = "all")

  if(!missing(savename)){
    save_as_docx(x, path = str_c(savename,".docx"))
    warning("Table has been save in directory")
  }
  return(x)
}




