###### Genotype × environment interaction and Stability analysis ##############

######################Stability analysis with metan ############################
############################### WAAS Model #####################################

library(metan)
library(ggplot2)
library(ggrepel)
library(writexl)
library(openxlsx)
options(max.print = 10000)

############################### data import ####################################

stabdata<-read.csv(file.choose(),)
attach(stabdata)
str(stabdata)
options(max.print = 10000)

######################### factors with unique levels ###########################

stabdata$ENV <- factor(stabdata$ENV, levels=unique(stabdata$ENV))
stabdata$GEN <- factor(stabdata$GEN, levels=unique(stabdata$GEN))
stabdata$REP <- factor(stabdata$REP, levels=unique(stabdata$REP))
str(stabdata)


###################### extract trait name from data file #######################

traitall <- colnames(stabdata)[sapply(stabdata, is.numeric)]
traitall

################### Data inspection and cleaning functions #####################

inspect(stabdata, threshold= 50, plot=FALSE) %>% rmarkdown::paged_table()


for (trait in traitall) {
  find_outliers(stabdata, var = all_of(trait), plots = TRUE)
}

remove_rows_na(stabdata)
replace_zero(stabdata)
find_text_in_num(stabdata, var = all_of(trait))

############################# data analysis ####################################
########################### descriptive stats ##################################

if (!file.exists("output")) {
  dir.create("output")
}

ds <- desc_stat(stabdata, stats="all", hist = TRUE, plot_theme = theme_metan())


write_xlsx(ds, file.path("output", "Descriptive.xlsx"))


############################# mean performances ################################
############################# mean of genotypes ################################

mg <- mean_by(stabdata, GEN) 
mg
View(mg)

############################# mean of environments #############################

me <- mean_by(stabdata, ENV)
me
View(me)

################################# two way mean #################################

dm <- mean_by(stabdata, GEN, ENV)
dm
View(dm)

############### mean performance of genotypes across environments ##############

mge <- stabdata %>% 
  group_by(ENV, GEN) %>%
  desc_stat(stats="mean")
mge
View(mge)

############### Exporting all mean performances computed above #################

write_xlsx(
  list(
    "Genmean" = mg,
    "Envmean" = me,
    "Genmeaninenv" = dm,
    "Genmeaninenv2" = mge
  ),
  file.path("output", "Mean Performance.xlsx")
)

############################ two-way table for all #############################

twgy_list <- list()

for (trait in traitall) {
  twgy_list[[trait]] <- make_mat(stabdata, GEN, ENV, val = trait)
}

result_twgy <- list()
for (trait in traitall) {
  twgy_result<- as.data.frame(twgy_list[[trait]])
  result_twgy[[trait]] <- twgy_result
}

twgy_wb <- createWorkbook()
for (trait in traitall) {
  addWorksheet(twgy_wb, sheetName = paste0(trait, ""))
  writeData(twgy_wb, sheet = trait, x = result_twgy[[trait]], startCol = 2, startRow = 1)
  writeData(twgy_wb, sheet = trait, x = rownames(twgy_list[[trait]]), startCol = 1, startRow = 2)
  writeData(twgy_wb, sheet = trait, x = c("GEN"), startCol = 1, startRow = 1)
}

saveWorkbook(twgy_wb, file.path("output", "TWmean.xlsx"), overwrite = TRUE)

################### plotting performance across environments ###################
#################### make performance for all traits in one ####################
################################## Heatmap #####################################

perfor_heat_list <- list()

for (trait in traitall) {
  perfor_heat_list[[trait]] <-
    ge_plot(
      stabdata,
      ENV,
      GEN,
      !!sym(trait),
      type = 1,
      values = FALSE,
      average = FALSE,
      text_col_pos = c("bottom"),
      text_row_pos = c("left"),
      width_bar = 1.5,
      heigth_bar = 20,
      xlab = "ENV",
      ylab = "GEN",
      plot_theme = theme_metan(),
      colour = TRUE
    ) + geom_tile(color = "transparent") + labs(title = paste0(trait, " performance across eight environments")) + theme(legend.title = element_text(), axis.text.x.bottom = element_text(angle = 0, hjust = .5)) + guides(fill = guide_colourbar(title = trait, barwidth = 1.5, barheight = 20))
  assign(paste0(trait, "_perfor_heat"), perfor_heat_list[[trait]])
}

############################## print all plots once ############################

for (trait in traitall) {
  print(perfor_heat_list[[trait]])
}

##################### high quality save all in one  ############################

if (!file.exists("output")) {
  dir.create("output")
}

if (!file.exists(file.path("output", "perfor_heat_plot"))) {
  dir.create(file.path("output", "perfor_heat_plot"))
}

for (trait in traitall) {
  ggsave(filename = file.path("output", "perfor_heat_plot", paste0(trait, ".png")),
         plot = perfor_heat_list[[trait]], width = 20, height = 30,
         dpi = 600, units = "cm")
}

################################# Line plot ####################################

perfor_line_list <- list()

for (trait in traitall) {
  perfor_line_list[[trait]] <-
    ge_plot(
      stabdata,
      ENV,
      GEN,
      !!sym(trait),
      type = 1,
      values = FALSE,
      average = FALSE,
      text_col_pos = c("bottom"),
      text_row_pos = c("left"),
      width_bar = 1.5,
      heigth_bar = 20,
      xlab = "ENV",
      ylab = "GEN",
      plot_theme = theme_metan(),
      colour = TRUE
    ) + geom_tile(color = "transparent") + labs(title = paste0(trait, " performance across eight environments")) + theme(legend.title = element_text(), axis.text.x.bottom = element_text(angle = 0, hjust = .5)) + guides(fill = guide_colourbar(title = trait, barwidth = 1.5, barheight = 20))
  assign(paste0(trait, "_perfor_line"), perfor_line_list[[trait]])
}

############################## print all plots once ############################

for (trait in traitall) {
  print(perfor_line_list[[trait]])
}

##################### high quality save all in one  ############################

if (!file.exists("output")) {
  dir.create("output")
}

if (!file.exists(file.path("output", "perfor_line_plot"))) {
  dir.create(file.path("output", "perfor_line_plot"))
}

for (trait in traitall) {
  ggsave(filename = file.path("output", "perfor_line_plot", paste0(trait, ".png")),
         plot = perfor_line_list[[trait]], width = 20, height = 30,
         dpi = 600, units = "cm")
}

########################## Genotype-environment winners ########################
#edit better argument as for some variables lower values are preferred and higher for others ####

traitall ### view your traits to decide for above condition ###

win <-
  ge_winners(
    stabdata,
    ENV,
    GEN,
    resp = everything(),
    better = c(h, h, h, h, h, h, h, h, l, l, l, h, h, h, h, h, h, h, l) 
  )
win
View(win)
ranks <-
  ge_winners(
    stabdata,
    ENV,
    GEN,
    resp = everything(),
    type = "ranks",
    better = c(h, h, h, h, h, h, h, h, l, l, l, h, h, h, h, h, h, h, l)
  )
ranks
View(ranks)

write_xlsx(list("winner" = win, "ranks" = ranks), file.path("output", "winner rank.xlsx"))

############################ ge or gge effects #################################
######################### combined for all ge effects ##########################
########################## ge effects to excel #################################

ge_list <- ge_effects(stabdata, ENV, GEN, resp = everything(), type = "ge")

result_ge_list <- list()
for (trait in traitall) {
  ge_list_result <- as.data.frame(ge_list[[trait]])
  result_ge_list[[trait]] <- ge_list_result
}

########################### save all in one excel  #############################

ge_list_wb <- createWorkbook()
for (trait in traitall) {
  addWorksheet(ge_list_wb, sheetName = paste0(trait, ""))
  writeData(ge_list_wb, sheet = trait, x = result_ge_list[[trait]])
}
saveWorkbook(ge_list_wb, file.path("output","ge_effects.xlsx"), overwrite = TRUE)

############################## ge effects plots ################################

ge_plots <- list()

for (trait in traitall) {
  ge_plots[[trait]] <- plot(ge_list) + aes(ENV, GEN) + theme(legend.title = element_text()) + guides(fill = guide_colourbar(title = paste0(trait, " ge effects"), barwidth = 1.5, barheight = 20))
}  ## also coord_flip() in place of aes

################# print all plots once #########################################

for (trait in traitall) {
  print(ge_plots[[trait]])
}

##################### high quality save all in one  ############################

if (!file.exists("output")) {
  dir.create("output")
}

if (!file.exists(file.path("output", "ge_plots"))) {
  dir.create(file.path("output", "ge_plots"))
}

for (trait in traitall) {
  ggsave(filename = file.path("output", "ge_plots", paste0(trait, ".png")),
         plot = ge_plots[[trait]], width = 20, height = 30,
         dpi = 600, units = "cm")
}

########################## gge effects to excel ################################

gge_list <- ge_effects(stabdata, ENV, GEN, resp = everything(), type = "gge")

result_gge_list <- list()
for (trait in traitall) {
  gge_list_result <- as.data.frame(gge_list[[trait]])
  result_gge_list[[trait]] <- gge_list_result
}

########################### save all in one excel  #############################

gge_list_wb <- createWorkbook()
for (trait in traitall) {
  addWorksheet(gge_list_wb, sheetName = paste0(trait, ""))
  writeData(gge_list_wb, sheet = trait, x = result_gge_list[[trait]])
}
saveWorkbook(gge_list_wb, file.path("output","gge_effects.xlsx"), overwrite = TRUE)

############################## ge effects plots ################################

gge_plots <- list()

for (trait in traitall) {
  gge_plots[[trait]] <- plot(gge_list) + aes(ENV, GEN) + theme(legend.title = element_text()) + guides(fill = guide_colourbar(title = paste0(trait, " gge effects"), barwidth = 1.5, barheight = 20))
}  ## also coord_flip() in place of aes

################# print all plots once #########################################

for (trait in traitall) {
  print(gge_plots[[trait]])
}

##################### high quality save all in one  ############################

if (!file.exists("output")) {
  dir.create("output")
}

if (!file.exists(file.path("output", "gge_plots"))) {
  dir.create(file.path("output", "gge_plots"))
}

for (trait in traitall) {
  ggsave(filename = file.path("output", "gge_plots", paste0(trait, ".png")),
         plot = gge_plots[[trait]], width = 20, height = 30,
         dpi = 600, units = "cm")
}

############################ fixed effect models ###############################
########################### Individual  and Joint anova ########################
######################### Individual anova for all traits ######################

aovind_list <- anova_ind(stabdata, env = ENV, gen = GEN, rep = REP, resp = everything())

result_aovind <- list()
for (trait in traitall) {
  ind_result<- as.data.frame(aovind_list[[trait]]$individual)
  result_aovind[[trait]] <- ind_result
}

aovind_wb <- createWorkbook()
for (trait in traitall) {
  addWorksheet(aovind_wb, sheetName = paste0(trait, ""))
  writeData(aovind_wb, sheet = trait, x = result_aovind[[trait]])
}
saveWorkbook(aovind_wb, file.path("output","indaovall.xlsx"), overwrite = TRUE)

################## Joint anova for all traits (ANOVA) ##########################

aovjoin_list <- anova_joint(stabdata, env = ENV, gen = GEN, rep = REP, resp = everything())

result_aovjoin <- list()
for (trait in traitall) {
  join_result<- as.data.frame(aovjoin_list[[trait]]$anova)
  result_aovjoin[[trait]] <- join_result
}

aovjoin_wb <- createWorkbook()
for (trait in traitall) {
  addWorksheet(aovjoin_wb, sheetName = paste0(trait, ""))
  writeData(aovjoin_wb, sheet = trait, x = result_aovjoin[[trait]])
}
saveWorkbook(aovjoin_wb, file.path("output","joinaovall.xlsx"), overwrite = TRUE)

################## Joint anova for all traits (Details) (2) ####################

result_aovjoin2 <- list()
for (trait in traitall) {
  join_result2<- as.data.frame(aovjoin_list[[trait]]$details)
  result_aovjoin2[[trait]] <- join_result2
}

aovjoin_wb2 <- createWorkbook()
for (trait in traitall) {
  addWorksheet(aovjoin_wb2, sheetName = paste0(trait, ""))
  writeData(aovjoin_wb2, sheet = trait, x = result_aovjoin2[[trait]])
}
saveWorkbook(aovjoin_wb2, file.path("output", "joinaovall2.xlsx"), overwrite = TRUE)

############### WAAS based stability analysis for all (ANOVA)###################

waas_list<-waas(stabdata, ENV, GEN, REP, resp = everything())

result_waas <- list()
for (trait in traitall) {
  waas_result<- as.data.frame(waas_list[[trait]]$ANOVA)
  result_waas[[trait]] <- waas_result
}

waas_wb <- createWorkbook()
for (trait in traitall) {
  addWorksheet(waas_wb, sheetName = paste0(trait, ""))
  writeData(waas_wb, sheet = trait, x = result_waas[[trait]])
}
saveWorkbook(waas_wb, file.path("output","waasanova.xlsx"), overwrite = TRUE)

############# WAAS based stability analysis for all (model) ####################

result_waas2 <- list()
for (trait in traitall) {
  waas_result2<- as.data.frame(waas_list[[trait]]$model)
  result_waas2[[trait]] <- waas_result2
}

waas_wb2 <- createWorkbook()
for (trait in traitall) {
  addWorksheet(waas_wb2, sheetName = paste0(trait, ""))
  writeData(waas_wb2, sheet = trait, x = result_waas2[[trait]])
}
saveWorkbook(waas_wb2, file.path("output", "waasmodel.xlsx"), overwrite = TRUE)

###################### WAAS based stability analysis for all (Means G*E)#####################################

result_waas3 <- list()
for (trait in traitall) {
  waas_result3<- as.data.frame(waas_list[[trait]]$MeansGxE)
  result_waas3[[trait]] <- waas_result3
}

waas_wb3 <- createWorkbook()
for (trait in traitall) {
  addWorksheet(waas_wb3, sheetName = paste0(trait, ""))
  writeData(waas_wb3, sheet = trait, x = result_waas3[[trait]])
}
saveWorkbook(waas_wb3, file.path("output","waasmeansge.xlsx"), overwrite = TRUE)

######################### WAAS biplots for all in one AMMI 1#################### 

waas1_list <- list()

for (trait in traitall) {
  waas1_list[[trait]] <- plot_scores(waas_list,
                                     var = trait,
                                     type = 1,
                                     first = "PC1",
                                     second = "PC2",
                                     x.lab = trait,
                                     repel = TRUE,
                                     max_overlaps = 50,
                                     shape.gen = 21,
                                     shape.env = 23,
                                     size.shape.gen = 2,
                                     size.shape.env = 3,
                                     col.bor.gen = "#215C29",
                                     col.bor.env = "#F68A31",
                                     col.line = "grey",
                                     col.gen = "#215C29",
                                     col.env = "#F68A31",
                                     col.segm.gen = transparent_color(),
                                     col.segm.env = "#F68A31",
                                     size.tex.gen = 3,
                                     size.tex.env = 3,
                                     size.tex.lab = 12,
                                     size.line = .4,
                                     line.type = 'dotdash',
                                     line.alpha = .8,
                                     highlight =,
                                     plot_theme = theme_metan(),
                                     size.segm.line = .4,
                                     leg.lab = c("Environment", "Genotype")) + labs(title = paste0("AMMI 1 biplot for ", trait)) + theme(
                                       plot.title = element_text(color = "black"),
                                       panel.background = element_rect(fill = "white"), plot.background = element_rect(fill = "white"),
                                       legend.position = c(0.88, .05),
                                       legend.background = element_rect(fill = NA)
                                     )
  assign(paste0(trait, "_waas1"), waas1_list[[trait]])
}
######## print all plots once #########################################

for (trait in traitall) {
  print(waas1_list[[trait]])
}

##################### high quality save all in one  ############################

if (!file.exists("output")) {
  dir.create("output")
}

if (!file.exists(file.path("output", "waas1_plots"))) {
  dir.create(file.path("output", "waas1_plots"))
}

for (trait in traitall) {
  ggsave(filename = file.path("output", "waas1_plots", paste0(trait, ".png")),
         plot = ammi1_list[[trait]], width = 15, height = 15,
         dpi = 600, units = "cm")
}

######################## Waas biplots for all in one AMMI 2 #################### 

waas2_list <- list()

for (trait in traitall) {
  waas2_list[[trait]] <- plot_scores(waas_list,
                                     var = trait,
                                     type = 2,
                                     first = "PC1",
                                     second = "PC2",
                                     repel = TRUE,
                                     max_overlaps = 50,
                                     shape.gen = 21,
                                     shape.env = 23,
                                     size.shape.gen = 2,
                                     size.shape.env = 3,
                                     col.bor.gen = "#215C29",
                                     col.bor.env = "#F68A31",
                                     col.line = "grey",
                                     col.gen = "#215C29",
                                     col.env = "#F68A31",
                                     col.segm.gen = transparent_color(),
                                     col.segm.env = "#F68A31",
                                     size.tex.gen = 3,
                                     size.tex.env = 3,
                                     size.tex.lab = 12,
                                     size.line = .4,
                                     line.type = 'dotdash',
                                     line.alpha = .8,
                                     highlight =,
                                     plot_theme = theme_metan(),
                                     size.segm.line = .4,
                                     leg.lab = c("Environment", "Genotype")) + labs(title = paste0("AMMI 2 biplot for ", trait)) + theme(
                                       plot.title = element_text(color = "black"),
                                       panel.background = element_rect(fill = "white"), plot.background = element_rect(fill = "white"),
                                       legend.position = c(0.88, .05),
                                       legend.background = element_rect(fill = NA))
  assign(paste0(trait, "_waas2"), waas2_list[[trait]])
}
######## print all plots once #########################################

for (trait in traitall) {
  print(waas2_list[[trait]])
}

##################### high quality save all in one  ############################

if (!file.exists("output")) {
  dir.create("output")
}

if (!file.exists(file.path("output", "waas2_plots"))) {
  dir.create(file.path("output", "waas2_plots"))
}

for (trait in traitall) {
  ggsave(filename = file.path("output", "waas2_plots", paste0(trait, ".png")),
         plot = ammi1_list[[trait]], width = 15, height = 15,
         dpi = 600, units = "cm")
}

############### Waas biplots for all in one (AMMI-Significant PCA) ############ 

waas3_list <- list()

for (trait in traitall) {
  waas3_list[[trait]] <- plot_scores(waas_list,
                                     var = trait,
                                     type = 3,
                                     first = "PC1",
                                     second = "PC2",
                                     x.lab = trait,
                                     repel = TRUE,
                                     max_overlaps = 50,
                                     shape.gen = 21,
                                     shape.env = 23,
                                     size.shape.gen = 2,
                                     size.shape.env = 3,
                                     col.bor.gen = "#215C29",
                                     col.bor.env = "#F68A31",
                                     col.line = "grey",
                                     col.gen = "#215C29",
                                     col.env = "#F68A31",
                                     col.segm.gen = transparent_color(),
                                     col.segm.env = "#F68A31",
                                     size.tex.gen = 3,
                                     size.tex.env = 3,
                                     size.tex.lab = 12,
                                     size.line = .4,
                                     line.type = 'dotdash',
                                     line.alpha = .8,
                                     highlight =,
                                     plot_theme = theme_metan(),
                                     size.segm.line = .4,
                                     leg.lab = c("Environment", "Genotype")) + labs(title = paste0("WAAS biplot for ", trait)) + theme(
                                       plot.title = element_text(color = "black"),
                                       panel.background = element_rect(fill = "white"), plot.background = element_rect(fill = "white"),
                                       legend.position = c(0.88, .05),
                                       legend.background = element_rect(fill = NA))
  assign(paste0(trait, "_waas3"), waas3_list[[trait]])
}

######## print all plots once #########################################

for (trait in traitall) {
  print(waas3_list[[trait]])
}

##################### high quality save all in one  ############################

if (!file.exists("output")) {
  dir.create("output")
}

if (!file.exists(file.path("output", "waas3_plots"))) {
  dir.create(file.path("output", "waas3_plots"))
}

for (trait in traitall) {
  ggsave(filename = file.path("output", "waas3_plots", paste0(trait, ".png")),
         plot = ammi1_list[[trait]], width = 15, height = 15,
         dpi = 600, units = "cm")
}

##################### WAAS scores plot for various traits ##################### 

waas_score_list <- list()

for (trait in traitall) {
  waas_score_list[[trait]] <- plot_waasby(waas_list,
                                     var = trait,
                                     size.shape = 3.5,
                                     size.tex.lab = 12,
                                     col.shape = c("#215C29", "#F68A31"),
                                     x.lab = "WAASY",
                                     y.lab = "Genotypes",
                                     x.breaks = waiver(),
                                     plot_theme = theme_metan()) + labs(title = paste0("WAASY score for ", trait)) + theme(
                                       plot.title = element_text(color = "black"),
                                       panel.background = element_rect(fill = "white"), plot.background = element_rect(fill = "white"),
                                       legend.position = c(0.88, .05),
                                       legend.background = element_rect(fill = NA))
  assign(paste0(trait, "_waas_score"), waas_score_list[[trait]])
}

######## print all plots once #########################################

for (trait in traitall) {
  print(waas_score_list[[trait]])
}

##################### high quality save all in one  ############################

if (!file.exists("output")) {
  dir.create("output")
}

if (!file.exists(file.path("output", "waas_score_plots"))) {
  dir.create(file.path("output", "waas_score_plots"))
}

for (trait in traitall) {
  ggsave(filename = file.path("output", "waas_score_plots", paste0(trait, ".png")),
         plot = ammi1_list[[trait]], width = 15, height = 15,
         dpi = 600, units = "cm")
}
