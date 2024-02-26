
#' TableFun
#'
#'@author Simon Pineau
#'
#'@description
#'Create an APA norme table.
#'
#' @param df A dataset (dataframe).
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
#' @param savename Optional argument. Saves the table in .docx format in the working space. The argument must be in quotes.
#'
#' @import rempsyc flextable export officer stringr dplyr tidyr
#' @return
#' @export
#' @examples
#' # Tableau "simple"
#' df<-data.frame("Variables" = c(1,2,3),
#'                "Mean" = c(32,22,15),
#'                "SD" = c(0.2,0.8,0.4))
#'
#' TableFun(df, header = "Mean and sd",foot.note.header = c(3,"a","SD = Standard deviation"))
#'
#' # Tableau avec plusieurs niveaux de noms de colonnes
#'
#' df<-data.frame("Variables" = c(1,2,3),
#'                "T1_Mean" = c(32,22,15),
#'                "T1_SD" = c(0.2,0.8,0.4),
#'                "T2_Mean" = c(30,24,18),
#'                "T2_SD" = c(0.3,0.5,0.6))
#'
#' TableFun(df,header = "Mean and sd across 2 times points", foot.note.header = c(3,"a","ET = Standard deviation"))
#'
#'
TableFun <- function(df, header, digits = 2, left.align = NULL,right.align = NULL, top.align = NULL, bottom.align = NULL, merge.col = NULL, first_col.header = NULL, footer = NULL, foot.note.header = NULL, foot.note.body = NULL, H.rotate = NULL, B.rotate = NULL, savename = NULL){
  x<-flextable(df) %>%
    separate_header() %>%
    autofit() %>%
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
    align(i = 1, part = "header", align = "left") %>%
    align(i = 1, part = "footer", align = "left")




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
    x <- rotate(x, i = c(-1),  j = c(H.rotate), rotation = "btlr", part = "header")
  }

  if(!missing(B.rotate)){
    x <- rotate(x, j = c(B.rotate), rotation = "btlr", part = "body")
  }

  x<-fix_border_issues(x,part = "all")

  if(!missing(first_col.header)){
    x<-as_grouped_data(x, groups = c(colnames(x[c(1)]))) %>%
      align(.,j = c(1), part = "body", align = "right") %>%
      merge_h(., i = c(1:2), part = "header")
  }



  if(!missing(savename)){
    save_as_docx(x, path = str_c(savename,".docx"))
    warning("Table has been save in directory")
  }
  return(x)
}
