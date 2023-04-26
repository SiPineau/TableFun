
#' TableFun
#'
#'@author Simon Pineau
#'
#'@description
#'Permet de créer des tableaux aux norme APA à partir de dataframe.
#'
#' @param df Un jeu de données (dataframe).
#' @param header Nom du tableau entre guillemets.
#' @param digits Nombre de décimale dans le tableau. Par default digits = 2.
#' @param left.align Argument optionnel. Aligne les informations du corps du tableau à gauche.
#' @param right.align Argument optionnel. Aligne les informations du corps du tableau à droite.
#' @param top.align Argument optionnel. Aligne les informations du corps du tableau en haut.
#' @param bottom.align Argument optionnel. Aligne les informations du corps du tableau en bas.
#' @param merge.col Funsionne les lignes similaires adjacentes des colonnes spécifiées. Par default merge.col = 1.
#' @param footer Argument optionnel. Insert une note de bas de tableau générale. L'argument doit être entre guillemets.
#' @param foot.note.header Argument optionnel. Insert une note specifique attachée un titre de colonne. Cette argument doit être un vecteur composé : du numéro de la colonne, de la "lettre" associé à la note, et de la "note specifique" : c(1,"a","note specifique"). Plusieurs notes spécifiques peuvent être ajoutées : c(1,"a","note specifique1", 2,"b","note specifique2").
#' @param foot.note.body Argument optionnel. Insert une note specifique attachée à une cellule du corps du tableau. La rédaction de l'argument est la même que celle de foot.note.header en ajoutant le numéro de la ligne après le numéro de la colonne : c(1, 1, "a","note specifique")
#' @param savename Argument optionnel. Enregistre le tableau au format .docx dans l'esapce de travail. L'argument doit être entre guillemets.
#'
#' @return
#' @export
#' @examples
#'
TableFun <- function(df, header, digits = 2, left.align = NULL,right.align = NULL, top.align = NULL, bottom.align = NULL, merge.col = 1, footer = NULL, foot.note.header = NULL, foot.note.body = NULL, savename = NULL){
  x<-flextable(df %>% mutate(across(where(is.numeric), round, digits))) %>%
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

    merge_v(j = merge.col, combine = TRUE) %>%
    align(part = "all", align = "center") %>%
    align(i = 1, part = "header", align = "left") %>%
    align(i = 1, part = "footer", align = "left")

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
    x<-align(x,j = c(right.align), part = "body", align = "left")
  }

  x<-fix_border_issues(x,part = "all")

  if(!missing(savename)){
    save_as_docx(x, path = str_c(savename,".docx"))
    warning("Table has been save in directory")
  }
  return(x)
}
