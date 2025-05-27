##### INTRODUCTION ################################################################################
# import input data
load("Input/InputReady.Mod2.Rdata")
source("Code/functions.needed.analysis.R")
source("Code/functions.needed.sci.not.R")
library(xlsx)
library(tidyverse)

soil.comp <- c("SoilMicro",
               "SoilMacro")

water.comp <- c("WaterMicro",
                "WaterMacro")

# product categories
PC <- unique(c(names(TC.Norm[["PP"]][["ProductSectorA"]]),
               names(TC.Norm[["PP"]][["ProductSectorB"]])))

# pre-consumer processes
comp.pre <- c("PrimaryProduction",
              "Transport")

# post-consumer processes
comp.post <- c("PCCollection",
               "PostCollection",
               "Recycling",
               "Incineration")

Mass.mat <- array(NA, c(length(Materials), length(Names), SIM),
                  dimnames = list(Materials, Names, NULL))

for(mat in Materials){
  load(paste0("Results/ResultsMass_NoSimplification/OutputMass_",mat,".Rdata"))
  for(comp in Names){
    Mass.mat[mat,comp,] <- Mass[comp,]  
  }
}

NiceNames <- as.matrix(read.xlsx(paste0("Input/",excel.file), sheetName = "Rank"))
rownames(NiceNames) <- NiceNames[,"Name"]

### SET UP VARIABLES ##############################################################################

# compartments after which to stop assessing emissions
PC.stop <- comp.post

# create a list to store the emissions
EF <- sapply(c("Pre-consumer processes", PC, "Post-consumer processes"), function(x) NULL)
for(i in 1:length(PC)){
  EF[[i]] <- sapply(Materials, function(x) NULL)
}

EF.mat <- array(0, c(4,length(Materials),length(EF),SIM),
                dimnames = list(c("Soil", "Soil (MP)", "Water", "Water (MP)"), Materials, names(EF), NULL))

### EMISSIONS FROM PRODUCTS #######################################################################

for(comp in PC){
  for(mat in Materials){
    EF[[comp]][[mat]] <- find.fc(comp = comp,
                                 mass = Mass.mat[mat,comp,],
                                 TC.Distr = TC.Norm[[mat]],
                                 verbose = F,
                                 stop.at = PC.stop)
    
    if(!is.null(EF[[comp]][[mat]][["SoilMicro"]])){
      EF.mat["Soil (MP)",mat,comp,] <- EF[[comp]][[mat]][["SoilMicro"]]
    }
    
    if(!is.null(EF[[comp]][[mat]][["SoilMacro"]])){
      EF.mat["Soil",mat,comp,] <- EF[[comp]][[mat]][["SoilMacro"]]
    }
    
    if(!is.null(EF[[comp]][[mat]][["WaterMicro"]])){
      EF.mat["Water (MP)",mat,comp,] <- EF[[comp]][[mat]][["WaterMicro"]]
    }
    
    if(!is.null(EF[[comp]][[mat]][["WaterMacro"]])){
      EF.mat["Water",mat,comp,] <- EF[[comp]][[mat]][["WaterMacro"]]
    }
    
  }
}

### EMISSIONS OUTSIDE OF CONSUMPTION ##############################################################

### PRODUCTION AND MANUFACTURING

for(comp in comp.pre){
  
  # remove the compartment of interest from the stop compartments
  comp.stop <- c(comp.pre, PC, comp.post)
  comp.stop <- comp.stop[-which(comp.stop == comp)]
  
  for(mat in Materials){
    
    EF[[comp]][[mat]] <- find.fc(comp = comp,
                                 mass = Mass.mat[mat,comp,],
                                 TC.Distr = TC.Norm[[mat]],
                                 verbose = F,
                                 stop.at = comp.stop)
    
    if(!is.null(EF[[comp]][[mat]][["SoilMicro"]])){
        EF.mat["Soil (MP)",mat,"Pre-consumer processes",] <- EF.mat["Soil (MP)",mat,"Pre-consumer processes",] + EF[[comp]][[mat]][["SoilMicro"]]
      }
    
    if(!is.null(EF[[comp]][[mat]][["SoilMacro"]])){
      EF.mat["Soil",mat,"Pre-consumer processes",] <- EF.mat["Soil",mat,"Pre-consumer processes",] + EF[[comp]][[mat]][["SoilMacro"]]
    }
    
    if(!is.null(EF[[comp]][[mat]][["WaterMicro"]])){
      EF.mat["Water (MP)",mat,"Pre-consumer processes",] <- EF.mat["Water (MP)",mat,"Pre-consumer processes",] + EF[[comp]][[mat]][["WaterMicro"]]
    }
    
    if(!is.null(EF[[comp]][[mat]][["WaterMacro"]])){
      EF.mat["Water",mat,"Pre-consumer processes",] <- EF.mat["Water",mat,"Pre-consumer processes",] + EF[[comp]][[mat]][["WaterMacro"]]
    }
    
  }
}


### END-OF-LIFE

for(comp in comp.post){
  
  # remove the compartment of interest from the stop compartments
  comp.stop <- comp.post
  comp.stop <- comp.stop[-which(comp.stop == comp)]
  
  for(mat in Materials){
    
    EF[[comp]][[mat]] <- find.fc(comp = comp,
                                 mass = Mass.mat[mat,comp,],
                                 TC.Distr = TC.Norm[[mat]],
                                 verbose = F,
                                 stop.at = comp.stop)
    
    if(!is.null(EF[[comp]][[mat]][["SoilMicro"]])){
      EF.mat["Soil (MP)",mat,"Post-consumer processes",] <- EF.mat["Soil (MP)",mat,"Post-consumer processes",] + EF[[comp]][[mat]][["SoilMicro"]]
    }
    
    if(!is.null(EF[[comp]][[mat]][["SoilMacro"]])){
      EF.mat["Soil",mat,"Post-consumer processes",] <- EF.mat["Soil",mat,"Post-consumer processes",] + EF[[comp]][[mat]][["SoilMacro"]]
    }
    
    if(!is.null(EF[[comp]][[mat]][["WaterMicro"]])){
      EF.mat["Water (MP)",mat,"Post-consumer processes",] <- EF.mat["Water (MP)",mat,"Post-consumer processes",] + EF[[comp]][[mat]][["WaterMicro"]]
    }
    
    if(!is.null(EF[[comp]][[mat]][["WaterMacro"]])){
      EF.mat["Water",mat,"Post-consumer processes",] <- EF.mat["Water",mat,"Post-consumer processes",] + EF[[comp]][[mat]][["WaterMacro"]]
    }
    
  }
}

### BAR PLOT ########################################################################

library(ggplot2)
library(sysfonts)
library(patchwork)

font_add(family = "Calibri", regular = "Calibri.ttf")

#color_vector <- c("#f8b862","#0094c8","#badcad","#fef263","#00a381","#f4b3c2","#a59aca")

## Soil MP

cp.order <- names(sort(apply(apply(EF.mat["Soil (MP)",,,],c(2,3),sum),1,mean)))
lab <- sci.not(apply(EF.mat["Soil (MP)",,,],3,sum)*1000)
labs <- cp.order <- cp.order[1:length(cp.order)]
labs[labs %in% PC] <- NiceNames[labs[labs %in% PC], "MediumLabel"]
mean.cal <- apply(EF.mat["Soil (MP)",,cp.order,],c(1,2),mean)*1000
sd.cal <- apply(EF.mat["Soil (MP)",,cp.order,],c(1,2),sd)*1000

colnames(mean.cal) <- labs
colnames(sd.cal) <- labs

data <- data.frame(
  polymer = rep(rownames(mean.cal), ncol(mean.cal)),
  prod = rep(colnames(mean.cal), each = nrow(mean.cal)),
  mean = as.vector(mean.cal),
  sd = as.vector(sd.cal)
)

mean_sum <- aggregate(mean ~ prod, data = data, FUN = sum)
sd_sum <- aggregate(sd ~ prod, data = data, FUN = sum)
data_sd <- merge(mean_sum, sd_sum)

# Reorder the dataframe

data$prod <- with(data,reorder(prod, mean, FUN = sum))
data$polymer <- factor(data$polymer, levels = c("LDPE", "HDPE", "PP", "PS", "EPS", "PVC", "PET"))

# Store results for output
data.to.print.1 <- data
data.to.print.1$comp <- rep('Soil MP',length(data$polymer))

# Bar plot

p1 <- ggplot(data, aes(x = prod, y = mean)) +
  geom_bar(aes(fill = polymer), stat = "identity", position = "stack", color = "black", width = 0.8) +  
  #scale_fill_manual(values = color_vector) + 
  scale_fill_brewer(palette="Set3")+
  theme_bw()+
  labs(x = "", y = "Mass (t)", fill = "Polymer") +
  ggtitle("Microplastic emissions to soil") +
  geom_pointrange(data = data_sd, aes(ymin = mean - sd, ymax = mean + sd), colour = "black", alpha = 0.9, size = 0.2) +
  ylim(0,NA) +
  coord_flip() +
  annotate ("text", x = -Inf, y = Inf, label = paste0("Σ = ",lab," t"), hjust = 1.05, vjust = -0.7)+
  theme(text = element_text(family = "Calibri"),
        axis.text = element_text(color = "black"))

p1

####### Soil MaP

cp.order <- names(sort(apply(apply(EF.mat["Soil",,,],c(2,3),sum),1,mean)))
lab <- sci.not(apply(EF.mat["Soil",,,],3,sum)*1000)
labs <- cp.order <- cp.order[1:length(cp.order)]
labs[labs %in% PC] <- NiceNames[labs[labs %in% PC], "MediumLabel"]
mean.cal <- apply(EF.mat["Soil",,cp.order,],c(1,2),mean)*1000
sd.cal <- apply(EF.mat["Soil",,cp.order,],c(1,2),sd)*1000

colnames(mean.cal) <- labs
colnames(sd.cal) <- labs

data <- data.frame(
  polymer = rep(rownames(mean.cal), ncol(mean.cal)),
  prod = rep(colnames(mean.cal), each = nrow(mean.cal)),
  mean = as.vector(mean.cal),
  sd = as.vector(sd.cal)
)

mean_sum <- aggregate(mean ~ prod, data = data, FUN = sum)
sd_sum <- aggregate(sd ~ prod, data = data, FUN = sum)
data_sd <- merge(mean_sum, sd_sum)

# Reorder the dataframe

data$prod <- with(data,reorder(prod, mean, FUN = sum))
data$polymer <- factor(data$polymer, levels = c("LDPE", "HDPE", "PP", "PS", "EPS", "PVC", "PET"))

# Store results for output
data.to.print.2 <- data
data.to.print.2$comp <- rep('Soil MaP',length(data$polymer))

# Bar plot

p2 <- ggplot(data, aes(x = prod, y = mean)) +
  geom_bar(aes(fill = polymer), stat = "identity", position = "stack", color = "black", width = 0.8) +  
  #scale_fill_manual(values = color_vector) + 
  scale_fill_brewer(palette="Set3")+
  theme_bw()+
  labs(x = "", y = "Mass (t)", fill = "Polymer") +
  ggtitle("Macroplastic emissions to soil") +
  geom_pointrange(data = data_sd, aes(ymin = mean - sd, ymax = mean + sd), colour = "black", alpha = 0.9, size = 0.2) +
  ylim(0,NA) +
  coord_flip() +
  annotate ("text", x = -Inf, y = Inf, label = paste0("Σ = ",lab," t"), hjust = 1.05, vjust = -0.7)+
  theme(text = element_text(family = "Calibri"),
        axis.text = element_text(color = "black"))

p2

####### Water MP

cp.order <- names(sort(apply(apply(EF.mat["Water (MP)",,,],c(2,3),sum),1,mean)))
lab <- sci.not(apply(EF.mat["Water (MP)",,,],3,sum)*1000)
labs <- cp.order <- cp.order[1:length(cp.order)]
labs[labs %in% PC] <- NiceNames[labs[labs %in% PC], "MediumLabel"]
mean.cal <- apply(EF.mat["Water (MP)",,cp.order,],c(1,2),mean)*1000
sd.cal <- apply(EF.mat["Water (MP)",,cp.order,],c(1,2),sd)*1000

colnames(mean.cal) <- labs
colnames(sd.cal) <- labs

data <- data.frame(
  polymer = rep(rownames(mean.cal), ncol(mean.cal)),
  prod = rep(colnames(mean.cal), each = nrow(mean.cal)),
  mean = as.vector(mean.cal),
  sd = as.vector(sd.cal)
)

mean_sum <- aggregate(mean ~ prod, data = data, FUN = sum)
sd_sum <- aggregate(sd ~ prod, data = data, FUN = sum)
data_sd <- merge(mean_sum, sd_sum)

# Avoid disappearing error bars
data_sd_setlimit <- data_sd %>%
  mutate(sd = if_else(data_sd$mean>data_sd$sd,sd,mean))

# Reorder the dataframe

data$prod <- with(data,reorder(prod, mean, FUN = sum))
data$polymer <- factor(data$polymer, levels = c("LDPE", "HDPE", "PP", "PS", "EPS", "PVC", "PET"))

# Store results for output
data.to.print.3 <- data
data.to.print.3$comp <- rep('Water MP',length(data$polymer))

# Bar plot

p3 <- ggplot(data, aes(x = prod, y = mean)) +
  geom_bar(aes(fill = polymer), stat = "identity", position = "stack", color = "black", width = 0.8) +  
  #scale_fill_manual(values = color_vector) + 
  scale_fill_brewer(palette="Set3")+
  theme_bw()+
  labs(x = "", y = "Mass (t)", fill = "Polymer") +
  ggtitle("Microplastic emissions to water") +
  geom_pointrange(data = data_sd_setlimit, aes(ymin = mean - sd, ymax = mean + sd), colour = "black", alpha = 0.9, size = 0.2) +
  ylim(0,NA) +
  coord_flip() +
  annotate ("text", x = -Inf, y = Inf, label = paste0("Σ = ",lab," t"), hjust = 1.05, vjust = -0.7)+
  theme(text = element_text(family = "Calibri"),
        axis.text = element_text(color = "black"))


p3

####### Water MaP

cp.order <- names(sort(apply(apply(EF.mat["Water",,,],c(2,3),sum),1,mean)))
lab <- sci.not(apply(EF.mat["Water",,,],3,sum)*1000)
labs <- cp.order <- cp.order[1:length(cp.order)]
labs[labs %in% PC] <- NiceNames[labs[labs %in% PC], "MediumLabel"]
mean.cal <- apply(EF.mat["Water",,cp.order,],c(1,2),mean)*1000
sd.cal <- apply(EF.mat["Water",,cp.order,],c(1,2),sd)*1000

colnames(mean.cal) <- labs
colnames(sd.cal) <- labs

data <- data.frame(
  polymer = rep(rownames(mean.cal), ncol(mean.cal)),
  prod = rep(colnames(mean.cal), each = nrow(mean.cal)),
  mean = as.vector(mean.cal),
  sd = as.vector(sd.cal)
)

mean_sum <- aggregate(mean ~ prod, data = data, FUN = sum)
sd_sum <- aggregate(sd ~ prod, data = data, FUN = sum)
data_sd <- merge(mean_sum, sd_sum)

# Reorder the dataframe

data$prod <- with(data,reorder(prod, mean, FUN = sum))
data$polymer <- factor(data$polymer, levels = c("LDPE", "HDPE", "PP", "PS", "EPS", "PVC", "PET"))

# Store results for output
data.to.print.4 <- data
data.to.print.4$comp <- rep('Water MaP',length(data$polymer))

# Avoid disappearing error bars
data_sd_setlimit <- data_sd %>%
  mutate(sd = if_else(data_sd$mean>data_sd$sd,sd,mean))

# Bar plot

p4 <- ggplot(data, aes(x = prod, y = mean)) +
  geom_bar(aes(fill = polymer), stat = "identity", position = "stack", color = "black", width = 0.8) +  
  #scale_fill_manual(values = color_vector) + 
  scale_fill_brewer(palette="Set3")+
  theme_bw()+
  labs(x = "", y = "Mass (t)", fill = "Polymer") +
  ggtitle("Macroplastic emissions to water") +
  geom_pointrange(data = data_sd_setlimit, aes(ymin = mean - sd, ymax = mean + sd), colour = "black", alpha = 0.9, size = 0.2) +
  coord_flip() +
  ylim(0,NA) +
  annotate ("text", x = -Inf, y = Inf, label = paste0("Σ = ",lab," t"), hjust = 1.05, vjust = -0.7)+
  theme(text = element_text(family = "Calibri"),
        axis.text = element_text(color = "black"))

p4

save <- p1 + p2 + p3 + p4 +
  plot_layout(guides = 'collect')

save

ggsave("EmissionsByProd.png",
       plot = save,
       path = "Results",
       width = 9.5, height = 6,
       dpi = 500)

data.to.print <- rbind(data.to.print.1,
                       data.to.print.2,
                       data.to.print.3,
                       data.to.print.4)
write.xlsx(data.to.print, "Results/EmissionsByProd.xlsx")
