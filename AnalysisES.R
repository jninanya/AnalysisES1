rm(list=ls())

#--- work directory ---#
WD <- "C:/Users/jninanya/OneDrive - CGIAR/Desktop/Prueba"
setwd(WD)

#--- reading libraries ---#
library(dplyr)
library(openxlsx)
library(agricolae)
#library(basicPlotteR) # To overlap labels
#library(shape) # To arrows plot

#--- reading the data ---#
#FileName <- "02_Data Matrix_Cutervo_2018.xlsx"
FileName <- "02_Data Matrix_Leopata_2018.xlsx"

df1 <- read.xlsx(FileName, sheet = "Matriz_Leopata")
df2 <- read.xlsx(FileName, sheet = "Organolepticos_HQ")

VL <- read.xlsx(FileName, sheet = "Var_List")

colnames(df1) <- tolower(colnames(df1))           # variables name in lower case
colnames(df2) <- tolower(colnames(df2))

#--- transformations ---#
df1$apgl_adj <- asin(sqrt(df1$apgl))*(180/3.1415) # asin is in radians
df1$tntp_adj <- sqrt(df1$tntp)

#--- qualitative (QL) and quantitative (QT) variables ---#
var <- VL[,c(3,4,6,7)]
colnames(var) <- c("var", "kind", "rep", "pca")

var$var <- tolower(var$var)
var$type[var$kind == "Cualitativa"] = "QL"
var$type[var$kind == "Cuantitativo"] = "QT"

QL1 <- var$var[var$type == "QL"]
QL <- QL1[-8]                                     # delete sppatt variable
QT <- var$var[var$type == "QT"]

#--- plot, row and colum as factor ---#
df1$plot <- as.factor(df1$plot)
df1$row <- as.factor(df1$row)
df1$colum <- as.factor(df1$colum)

#--- function: number of data (nd) ---#
nd <- function(x){
  x = x[!is.na(x)]
  nd = length(x)
}

#--- function: coefficient of variation (CV) ---#
CV <- function(x, na.rm = FALSE){
  sd(x, na.rm = na.rm)/mean(x, na.rm = na.rm)
}
 
#--- summary by instn  for df1 ---#
sm1_nd <- df1 %>%
  group_by(instn) %>%
  summarise_if(is.numeric, nd)

sm1_mean <- df1 %>%
  group_by(instn) %>%
  summarise_if(is.numeric, mean, na.rm = TRUE)

sm1_sd <- df1 %>%
  group_by(instn) %>%
  summarise_if(is.numeric, sd, na.rm = TRUE)

sm1_median <- df1 %>%
  group_by(instn) %>%
  summarise_if(is.numeric, median, na.rm = TRUE)

#--- summary by instn  for df2 ---#
sm2_nd <- df2 %>%
  group_by(instn) %>%
  summarise_if(is.numeric, nd)

sm2_mean <- df2 %>%
  group_by(instn) %>%
  summarise_if(is.numeric, mean, na.rm = TRUE)

sm2_sd <- df2 %>%
  group_by(instn) %>%
  summarise_if(is.numeric, sd, na.rm = TRUE)

sm2_median <- df2 %>%
  group_by(instn) %>%
  summarise_if(is.numeric, median, na.rm = TRUE)

#--- dplyr to data.frame ---#
sm1_nd <- as.data.frame(sm1_nd)
sm1_mean <- as.data.frame(sm1_mean)
sm1_sd <- as.data.frame(sm1_sd)
sm1_median <- as.data.frame(sm1_median)

sm2_nd <- as.data.frame(sm2_nd)
sm2_mean <- as.data.frame(sm2_mean)
sm2_sd <- as.data.frame(sm2_sd)
sm2_median <- as.data.frame(sm2_median)

#--- export summary for QL variables ---#
dfr_QL_nd <- data.frame(sm1_nd[, QL[1:8]], sm2_nd[, QL[9:12]])
dfr_QL_mean <- data.frame(sm1_mean[, QL[1:8]], sm2_mean[, QL[9:12]])
dfr_QL_median <- data.frame(sm1_median[, QL[1:8]], sm2_median[, QL[9:12]])

nQL <- length(QL)
mtx_QL <- matrix(nrow = nrow(sm1_nd), ncol = (2*nQL))
nameQL <- vector()

iQL <- QL#[order(QL)]

for(i in 1:nQL){
  mtx_QL[,1+2*(i-1)] <- dfr_QL_nd[, iQL[i]] 
  mtx_QL[,2+2*(i-1)] <- dfr_QL_median[, iQL[i]]
#  mtx_QL[,3+3*(i-1)] <- dfr_QL_median[, iQL[i]]

  nameQL[1+2*(i-1)] <- paste0(iQL[i], "_n")  
  nameQL[2+2*(i-1)] <- paste0(iQL[i], "_median")
#  nameQL[3+3*(i-1)] <- paste0(iQL[i], "_median")
}

dfr_QL <- data.frame(sm1_nd$instn, mtx_QL)
colnames(dfr_QL) <- c("instn", nameQL)

write.csv(dfr_QL, "summary_of_QL_traits.csv")             # export in csv format

#--- export summary for QL variables ---#
nQT <- length(QT)
mtx_QT <- matrix(nrow = nrow(sm1_nd), ncol = (3*nQT))
nameQT <- vector()

iQT <- QT[order(QT)]

for(i in 1:nQT){
  mtx_QT[,1+3*(i-1)] <- sm1_nd[, iQT[i]]
  mtx_QT[,2+3*(i-1)] <- sm1_mean[, iQT[i]] 
  mtx_QT[,3+3*(i-1)] <- sm1_sd[, iQT[i]]
#  mtx_QT[,4+4*(i-1)] <- sm1_median[, QT[i]]
  
  nameQT[1+3*(i-1)] <- paste0(iQT[i], "_n")
  nameQT[2+3*(i-1)] <- paste0(iQT[i], "_mean")
  nameQT[3+3*(i-1)] <- paste0(iQT[i], "_sd")
#  nameQT[4+4*(i-1)] <- paste0(QT[i], "_median")
}

dfr_QT <- data.frame(sm1_mean$instn, mtx_QT)
colnames(dfr_QT) <- c("instn", nameQT)

write.csv(dfr_QT, "summary_of_QT_traits.csv")            # export in csv format

#--- final matriz ---#
dfr_f <- data.frame(sm1_mean[, c("instn", QT)], sm1_median[, QL[1:8]], sm2_median[, QL[9:12]])

####################################
####################################

#--- ANOVA ---#
name_ix <- var$var[var$type == "QT" & var$rep == "3 rep"]
dir.create(paste0(WD, "/TukeyTest/"), showWarnings = FALSE)
dir.create(paste0(WD, "/TukeyTest/ANOVA"), showWarnings = FALSE)
dir.create(paste0(WD, "/TukeyTest/Histograms"), showWarnings = FALSE)
dir.create(paste0(WD, "/TukeyTest/QQPlots"), showWarnings = FALSE)

r_adj <- vector()
r_squared <- vector()
p_value <- vector()

for(i in 1:length(name_ix)){
ix <- df1[, c("instn", "rep")]
ix$trt <- df1[, name_ix[i]]

model <- lm(trt ~ instn + rep, data = ix)
ao <- anova(model)
fm <- aov(model)
p <- summary(model)

out <- LSD.test(model,"instn", p.adj="bonferroni")
#out2 <- HSD.test(model, "instn")

rs <- out$groups
colnames(rs) <- c(name_ix[i], "group")

r_squared[i] <- p$r.squared
r_adj[i] <- p$adj.r.squared

write.csv(ao, paste0(WD,"/TukeyTest/ANOVA/anova_for_", name_ix[i],".csv"))
write.csv(rs, paste0(WD,"/TukeyTest/tukey_test_results_", name_ix[i],".csv"))

# Validacion del modelo
F_Hist <- paste0(WD,"/TukeyTest/Histograms/Histogram_", name_ix[i], ".png")            # for QQ plot
png(file.path(F_Hist), width=800, height=600, units="px", pointsize=12)
 hist(fm$residuals, main = paste0("Histogram of ", name_ix[i]),
      xlab = "residuals", breaks = 8)
dev.off()

F_QQPlot <- paste0(WD,"/TukeyTest/QQPlots/QQPlot_", name_ix[i], ".png")            # for QQ plot
png(file.path(F_QQPlot), width=800, height=600, units="px", pointsize=12)
 qqnorm(fm$residuals, main = paste0("Normal Q-Q Plot of ", name_ix[i], "'s residuals"))
 qqline(fm$residuals, col = "red1")
dev.off()


st <- shapiro.test(fm$residuals)
p_value[i] <- st$p.value

}

sh_t <- data.frame("var" = name_ix, "p_value" = round(p_value, 3))
sh_t$is.normal <- ifelse(sh_t$p_value < 0.05, FALSE, TRUE)
sh_t <- sh_t[order(sh_t$var),]

write.csv(sh_t, paste0(WD, "/TukeyTest/shapiro_test_results.csv"))


#####################
p_model <- data.frame(name_ix, r_squared, r_adj)
p_model <- p_model[order(p_model$name_ix), ]
colnames(p_model) <- c("var", "r_squared", "r_squared_adj")

write.csv(p_model, paste0(WD,"/TukeyTest/model's_r_squared.csv"))


### KruskalTest and Friedman 
#dir.create(paste0(WD, "/KruskalTest/"), showWarnings = FALSE)
dir.create(paste0(WD, "/FriedmanTest/"), showWarnings = FALSE)
name_QL <- QL[1:8]

for(i in 1:length(name_QL)){

  ix <- df1[, c("instn", "rep")]
  ix$trait <- df1[, name_QL[i]]
  dat = na.omit(ix)
  
  k <- friedman(judge = dat$rep, trt = dat$instn, evaluation = dat$trait, alpha = 0.05, group = TRUE, main = NULL)
  #k = (kruskal(y = dat$trait, trt = dat$instn, p.adj = "bonferroni"))
  
  colnames(k$means)[1] <- name_QL[i]
  colnames(k$groups)[1] <- paste0(name_QL[i], "_rank")
  
#  write.csv(k$means, paste0(WD, "/KruskalTest/means_kruskal_test_", name_QL[i], ".csv"))
#  write.csv(k$groups, paste0(WD, "/KruskalTest/groups_kruskal_test_", name_QL[i], ".csv"))
  
  write.csv(k$statistics, paste0(WD, "/FriedmanTest/statistics_friedman_test_", name_QL[i], ".csv"))
  write.csv(k$means, paste0(WD, "/FriedmanTest/means_friedman_test_", name_QL[i], ".csv"))
  write.csv(k$groups, paste0(WD, "/FriedmanTest/groups_friedman_test_", name_QL[i], ".csv"))
}

#--- correlation matriz ---#
mtx_cor <- cor(dfr_f[,c(QT, QL)],  use = "pairwise.complete.obs")
write.csv(mtx_cor, "correlation_matrix.csv")

#--- overall summary ---#
vf_QT <- data.frame("var_QT" = QT)
vf_QL <- data.frame("var_QL" = QL)

for(i in 1:nQT){
  vf_QT$mean[i] <- mean(dfr_f[,QT[i]], na.rm = TRUE)
  vf_QT$median[i] <- median(dfr_f[,QT[i]], na.rm = TRUE)
  vf_QT$sd[i] <- sd(dfr_f[,QT[i]], na.rm = TRUE)
  vf_QT$cv[i] <- CV(dfr_f[,QT[i]], na.rm = TRUE)
  vf_QT$max[i] <- max(dfr_f[,QT[i]], na.rm = TRUE)
  vf_QT$min[i] <- min(dfr_f[,QT[i]], na.rm = TRUE)
}

for(i in 1:nQL){
  vf_QL$median <- median(dfr_f[, QL[i]], na.rm = TRUE)
#  vf2$mode <- mfv(dfr_f[, QL[i]], na_rm = TRUE)
}

write.csv(vf_QT, "descriptive_QT.csv")
write.csv(vf_QL, "descriptive_QL.csv")

#--- principal componente analysis ---#
dir.create(paste0(WD, "/PCA/"), showWarnings = FALSE)
# for org traits 
ot <- dfr_f[, QL[9:12]]
rownames(ot) <- dfr_f$instn

res.pca <- prcomp(ot, center = TRUE, scale = TRUE)
sum.pca <- summary(res.pca)

write.csv(res.pca$rotation, paste0(WD, "/PCA/pca_for_og_traits_var_comp.csv"))
write.csv(res.pca$x, paste0(WD, "/PCA/pca_for_og_traits_ind_comp.csv"))
write.csv(sum.pca$importance, paste0(WD, "/PCA/pca_for_og_traits_eigenvalues.csv"))

# for rest traits
traits <- var$var[var$pca == "si"]

ot2 <- dfr_f[, traits]
rownames(ot2) <- dfr_f$instn

res.pca2 <- prcomp(ot2, center = TRUE, scale = TRUE)
sum.pca2 <- summary(res.pca2)

write.csv(res.pca2$rotation, paste0(WD, "/PCA/pca_for_selec_traits_var_comp.csv"))
write.csv(res.pca2$x, paste0(WD, "/PCA/pca_for_selec_traits_ind_comp.csv"))
write.csv(sum.pca2$importance, paste0(WD, "/PCA/pca_for_selec_traits_eigenvalues.csv"))

###