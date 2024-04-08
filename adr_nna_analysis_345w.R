#AREA OF RESPONSABILITY CHILD PROTECTION VENEZUELA 2024
# INDICATORS
# SECTORIAL OBJECTIVES
# REACHED PEOPLE

rm(list=ls())
if (!require("pacman")) install.packages("pacman")
pacman::p_load(tidyverse, readxl, openxlsx, writexl, sf, leaflet, fuzzyjoin, dplyr)
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
getwd()
source("utils.R")
# START -------------------------------------------------------------------

folder <- getwd()
file1 <- "345WDataset_MD87.xlsx"
data <- read.xlsx(xlsxFile = paste0(folder,'/01_input/',file1), sheet ='Hoja-1', startRow = 2)

folder <- getwd()
file2 <- "2024_adr_nna_indicadores.xlsx"
indicators <- read.xlsx(xlsxFile = paste0(folder,'/01_input/',file2), sheet ='nna', startRow = 1, fillMergedCells = TRUE, colNames = TRUE)


# EXPORT FILES ------------------------------------------------------------
output <- "2024_02_adr_nna_analisis_5W"
output.ocha <- "2024_02_nna_5w_ocha"
col_width <- 20


# ORGANISE DATA -----------------------------------------------------------
# Organise selection of activities by indicator
indicators<-indicators %>% 
  mutate(number_activity = sub("([^:]*).*", "\\1", activity))
indicators.list <- unique(indicators[c('indicator')])


# Copy raw data
df<-data

# Remove special character or accent
df[col.validated]<-iconv(df[[col.validated]],from="UTF-8",to="ASCII//TRANSLIT")

df<-df %>% 
  # filter(tolower(!!sym(col.validated)) == 'si')
filter(tolower(!!sym(col.recurrent)) == 'no',
       tolower(!!sym(col.validated)) == 'si')

# Convert columns to numeric class
df<-df %>%
  mutate(across(all_of(c(col.reach.type, col.reach.disagg)), as.numeric)) %>% 
  mutate(across(all_of(c(col.reach.type, col.reach.disagg)), ~ ifelse(is.nan(.)| is.na(.), 0, .)))


# Indicador 1.1.2.4 -------------------------------------------------------------
# de NNA y personas cuidadoras afectados y en riesgo con acceso a actividades de salud mental y apoyo psicosocial individual y grupal utilizando un enfoque diferencial de género, edad y diversidad
# max(2.13,2.14)+ max(2.15,2.16)

name.act <- unlist((indicators %>% filter(indicator == indicators.list[1, "indicator"]) %>% select(number_activity)))
name.ind <- unique(unlist((indicators %>% filter(indicator == indicators.list[1, "indicator"]) %>% select(indicator))))

name.max <- c("2.13","2.14")
temp1<-calculate_max(data = df,activity = name.max,aggregation = c(adm.level[[3]],location, col.activity),targets_summarise = c(col.reach.type, col.reach.disagg))

name.max <-c("2.15","2.16")
temp2<-calculate_max(data = df,activity = name.max,aggregation = c(adm.level[[3]],location, col.activity),targets_summarise = c(col.reach.type, col.reach.disagg))

temp <- bind_rows(temp1, temp2)
temp <- calculate_sum(data = temp,
                      activity = name.act,
                      aggregation = adm.level[[3]],
                      targets_summarise = c(col.reach.type,col.reach.disagg)) 


ind.1124 <- temp %>% 
  mutate(indicator = unique(name.ind),.before = 1) 


# Indicador 1.1.2.5 -------------------------------------------------------------
# de NNA afectados y en riesgos de protección que acceden a servicios de protección de la niñez utilizando un enfoque de género, edad y diversidad
# sum(2.17, 2.18, 2.19)

name.act <- unlist((indicators %>% filter(indicator == indicators.list[2, "indicator"]) %>% select(number_activity)))
name.ind <- unique(unlist((indicators %>% filter(indicator == indicators.list[2, "indicator"]) %>% select(indicator))))

temp <- calculate_sum(data = df,
                      activity = name.act,
                      aggregation = c(adm.level[[3]]),
                      targets_summarise = c(col.reach.type, col.reach.disagg))


ind.1125 <- temp %>% 
  mutate(indicator = unique(name.ind),.before = 1) 


# INDICATOR 2.2.1.3 -----------------------------------------------------------
# 2.2.1.3 # de NNA que acceden al registro civil de nacimientos y otros docuimentos de identidad
# 4.03 + 4.04

name.act <- unlist((indicators %>% filter(indicator == indicators.list[3, "indicator"]) %>% select(number_activity)))
name.ind <- unique(unlist((indicators %>% filter(indicator == indicators.list[3, "indicator"]) %>% select(indicator))))

temp <- calculate_sum(data = df,
                      activity = name.act,
                      aggregation = c(adm.level[[3]]),
                      targets_summarise = c(col.reach.type, col.reach.disagg))


ind.2213 <- temp %>% 
  mutate(indicator = unique(name.ind),.before = 1) 


# INDICATOR 3.3.1.1 -----------------------------------------------------------
# 3.3.1.1: # de personas de la comunidad capacitadas y sensibilizadas en temas de protección de niños, niñas y adolescentes 
# sum(10.01, 10.02)
# remove virtual activities for the calculation of reached people, but not for indicator results

name.act <- unlist((indicators %>% filter(indicator == indicators.list[4, "indicator"]) %>% select(number_activity)))
name.ind <- unique(unlist((indicators %>% filter(indicator == indicators.list[4, "indicator"]) %>% select(indicator))))

temp.virtual <- calculate_sum(data = df,
                              activity=name.act,
                              aggregation = c(adm.level[[3]]),
                              targets_summarise = c(col.reach.type, col.reach.disagg))

temp <- calculate_sum(df %>% filter(tolower(!!sym(col.loc.type))!="virtual"),
                              activity=name.act,
                              aggregation = c(adm.level[[3]]),
                      targets_summarise = c(col.reach.type, col.reach.disagg))


ind.3311.virtual <- temp.virtual %>% 
  mutate(indicator = unique(name.ind),.before = 1)   


ind.3311 <- temp %>% 
  mutate(indicator = unique(name.ind),.before = 1)  


# INDICATOR 3.3.1.2 -----------------------------------------------------------
# 3.3.1.2: # de redes comunitarias creadas y/o fortalecidas para la protección de NNA
# 10.3

name.act <- unlist((indicators %>% filter(indicator == indicators.list[5, "indicator"]) %>% select(number_activity)))
name.ind <- unique(unlist((indicators %>% filter(indicator == indicators.list[5, "indicator"]) %>% select(indicator))))

temp <- calculate_sum(data = df,
                      activity = name.act,
                      aggregation = c(adm.level[[3]]),
                      targets_summarise =  c(col.reach.type, col.reach.disagg))


ind.3312 <- temp %>% 
  mutate(indicator = unique(name.ind),.before = 1)  


# INDICATOR 3.3.2.1 -----------------------------------------------------------
# Indicador 3.3.2.1: # de personas de instituciones del Estado y sociedad civil capacitadas y apoyadas con asistencia técnica en temas de protección de niños, niñas y adolescentes 
# sum(1.01, 1.02)

name.act <- unlist((indicators %>% filter(indicator == indicators.list[6, "indicator"]) %>% select(number_activity)))
name.ind <- unique(unlist((indicators %>% filter(indicator == indicators.list[6, "indicator"]) %>% select(indicator))))

max.act <- c('1.01')
temp.max1 <- calculate_max(data = df,
                      activity = max.act,
                      aggregation = c(adm.level[[3]], location),
                      targets_summarise =  c(col.reach.type, col.reach.disagg)) 

max.act <- c('1.02')
temp.max2 <- calculate_max(data = df,
                           activity = max.act,
                           aggregation = c(adm.level[[3]], location),
                           targets_summarise =  c(col.reach.type, col.reach.disagg)) 

temp <- bind_rows(temp.max1, temp.max2)

temp <- calculate_sum(data = temp,
                      activity = 'none' ,
                      aggregation = c(adm.level[[3]]),
                      targets_summarise = c(col.reach.type, col.reach.disagg))

ind.3321 <- temp %>% 
  mutate(indicator = unique(name.ind),.before = 1)  


# INDICATOR 3.3.2.2 -----------------------------------------------------------
# 3.3.2.2: # de personas que acceden a los servicios de protección de la niñez de instituciones apoyadas con dotaciones
# sum(1.03)

name.act <- unlist((indicators %>% filter(indicator == indicators.list[7, "indicator"]) %>% select(number_activity)))
name.ind <- unique(unlist((indicators %>% filter(indicator == indicators.list[7, "indicator"]) %>% select(indicator))))


temp <- calculate_sum(data = df,
                      activity = name.act,
                      aggregation = c(adm.level[[3]]),
                      targets_summarise = c(col.reach.type, col.reach.disagg))

ind.3322 <- temp %>% 
  mutate(indicator = unique(name.ind),.before = 1) 



# INDICATOR TOTAL ---------------------------------------------------------

temp <- bind_rows(ind.1124,ind.1125, ind.2213, ind.3311.virtual, ind.3312, ind.3321, ind.3322)

ind.total <- calculate_sum(data = temp,
                           activity = 'none',
                           aggregation = 'indicator',
                           targets_summarise = c(col.reach.type,col.reach.disagg))

# Calculate disaggregation by age and sex (ninas, ninos, mujeres, homres, mujeres mayor y hombres mayor)
ind.total<-calculate_age_gender_disaggregation(ind.total)
ind.1124<-calculate_age_gender_disaggregation(ind.1124)
ind.1125<-calculate_age_gender_disaggregation(ind.1125)
ind.2213<-calculate_age_gender_disaggregation(ind.2213)
ind.3311.virtual<-calculate_age_gender_disaggregation(ind.3311.virtual)
ind.3311<-calculate_age_gender_disaggregation(ind.3311)
ind.3312<-calculate_age_gender_disaggregation(ind.3312)
ind.3321<-calculate_age_gender_disaggregation(ind.3321)
ind.3322<-calculate_age_gender_disaggregation(ind.3322)

# Calculate indicators by administrative level
temp <- bind_rows(ind.1124,ind.1125, ind.2213, ind.3311.virtual, ind.3312, ind.3321, ind.3322)

ind.adm3 <- calculate_sum(data = temp ,activity = 'none',aggregation = c('indicator', adm.level[[3]]), targets_summarise =c(col.reach.type,col.reach.disagg, col.totals, col.soma))
ind.adm2 <- calculate_sum(data = ind.adm3 ,activity = 'none',aggregation = c('indicator',adm.level[[2]]), targets_summarise =c(col.reach.type,col.reach.disagg,col.totals,col.soma))
ind.adm1 <- calculate_sum(data = ind.adm3 ,activity = 'none',aggregation = c('indicator', adm.level[[1]]), targets_summarise =c(col.reach.type,col.reach.disagg,col.totals,col.soma))
ind.adm0 <- calculate_sum(data = ind.adm3 ,activity = 'none',aggregation = c('indicator'), targets_summarise =c(col.reach.type,col.reach.disagg,col.totals,col.soma))


# SECTORIAL OBJECTIVES ----------------------------------------------------

objectives.list<-as.vector(unique(indicators$sectorial_objective) %>% na.omit())
oe.list <- lapply(1:length(objectives.list), function(i){
  name.ind<-unique(unlist((indicators %>% filter(sectorial_objective == objectives.list[i]) %>% select(indicator))))
  
  temp <- filter(ind.adm3, grepl(paste(name.ind, collapse = "|"),indicator))
  oe.adm3 <- calculate_sum(data = temp, activity = 'none', aggregation = adm.level[[3]], targets_summarise =c(col.reach.type,col.reach.disagg,col.totals,col.soma)) %>%
    mutate(indicator = objectives.list[[i]], .before = 1)
  
  oe.adm2 <- calculate_sum(data = oe.adm3, activity = 'none', aggregation = c('indicator', adm.level[[2]]), targets_summarise =c(col.reach.type,col.reach.disagg,col.totals,col.soma))
  oe.adm1 <- calculate_sum(data = oe.adm3, activity = 'none', aggregation = c('indicator', adm.level[[1]]), targets_summarise =c(col.reach.type,col.reach.disagg,col.totals,col.soma))
  oe.adm0 <- calculate_sum(data = oe.adm3, activity = 'none', aggregation = c('indicator'), targets_summarise =c(col.reach.type,col.reach.disagg,col.totals,col.soma))
  result <- list(oe.adm0, oe.adm1, oe.adm2, oe.adm3)
})

oe.adm0 <- bind_rows(oe.list[[1]][[1]], oe.list[[2]][[1]], oe.list[[3]][[1]])
oe.adm1 <- bind_rows(oe.list[[1]][[2]], oe.list[[2]][[2]], oe.list[[3]][[2]])
oe.adm2 <- bind_rows(oe.list[[1]][[3]], oe.list[[2]][[3]], oe.list[[3]][[3]])
oe.adm3 <- bind_rows(oe.list[[1]][[4]], oe.list[[2]][[4]], oe.list[[3]][[4]])


# TOTAL ALCANZADOS --------------------------------------------------------

temp <- bind_rows(ind.1124,ind.1125, ind.2213, ind.3311, ind.3321)

# Create dataframe with adults and children
target.adm3 <- calculate_sum(data = temp, activity = 'none',aggregation = adm.level[[3]],targets_summarise = c(col.reach.type,col.reach.disagg, col.totals, col.soma))
target.adm2 <- calculate_sum(data = temp, activity = 'none',aggregation = adm.level[[2]],targets_summarise = c(col.reach.type,col.reach.disagg, col.totals, col.soma))
target.adm1 <- calculate_sum(data = temp, activity = 'none',aggregation = adm.level[[1]],targets_summarise = c(col.reach.type,col.reach.disagg, col.totals, col.soma))
target.adm0 <- calculate_sum(data = temp %>% mutate('#adm0+name'='Venezuela', '#adm0+code'='VEN0'),activity = 'none',aggregation = c('#adm0+name','#adm0+code'),targets_summarise = c(col.reach.type,col.reach.disagg, col.totals,col.soma))


# CALCULATE PERCENTAGE ----------------------------------------------------
# Calculate Percentages
# Add columns with percentage values


# Calculate percentage of indicator and objective
df.adm0 <- calculate_percentage(arrange_data(bind_rows(oe.adm0, ind.adm0), adm.level = 'adm0'), adm.level = 'adm0') %>%
  filter(!desagregacion %in% c('#reached+f+adults', '#reached+m+adults', '#reached+f+elderly', '#reached+m+elderly')) %>% 
  mutate('#adm0+name'='Venezuela', '#adm0+code'='VEN0', .before=1)
df.adm1 <- calculate_percentage(arrange_data(bind_rows(oe.adm1, ind.adm1), adm.level = adm1), adm.level = adm.level[[1]]) %>% filter(!desagregacion %in% c('#reached+f+adults', '#reached+m+adults', '#reached+f+elderly', '#reached+m+elderly')) %>% 
  select(all_of(adm.level[[1]]), everything())
df.adm2 <- calculate_percentage(arrange_data(bind_rows(oe.adm2, ind.adm2), adm.level = adm12), adm.level = adm.level[[2]]) %>% filter(!desagregacion %in% c('#reached+f+adults', '#reached+m+adults', '#reached+f+elderly', '#reached+m+elderly')) %>% 
  select(all_of(adm.level[[2]]), everything())
# df.adm3 <- calculate_percentage(arrange_data(bind_rows(oe.adm3, ind.adm3), adm.level = adm123), adm.level = adm.level[[3]]) %>% filter(!desagregacion %in% c('#reached+f+adults', '#reached+m+adults', '#reached+f+elderly', '#reached+m+elderly'))


t.adm0 <- calculate_percentage_target(arrange_data_target(target.adm0,adm.level = adm0), adm.level = adm0) %>% filter(!desagregacion %in% c('#reached+f+adults', '#reached+m+adults', '#reached+f+elderly', '#reached+m+elderly'))
t.adm1 <- calculate_percentage_target(arrange_data_target(target.adm1,adm.level = adm1), adm.level = adm.level[[1]]) %>% filter(!desagregacion %in% c('#reached+f+adults', '#reached+m+adults', '#reached+f+elderly', '#reached+m+elderly'))
t.adm2 <- calculate_percentage_target(arrange_data_target(target.adm2,adm.level = adm12), adm.level = adm.level[[2]]) %>% filter(!desagregacion %in% c('#reached+f+adults', '#reached+m+adults', '#reached+f+elderly', '#reached+m+elderly'))
# t.adm3 <- calculate_percentage_target(arrange_data_target(target.adm3,adm.level = adm123), adm.level = adm.level[[3]]) %>% filter(!desagregacion %in% c('#reached+f+adults', '#reached+m+adults', '#reached+f+elderly', '#reached+m+elderly'))


# EXPORT ------------------------------------------------------------------

# Create list of dataframes
df_list <- list(
  # 'DATA' = df,
  'adm0_target' = t.adm0,
  'adm0_ind_obj' = df.adm0,
  
  'adm1_target' = t.adm1,
  'adm1_ind_obj' = df.adm1,
  
  'adm2_target' = t.adm2,
  'adm2_ind_obj' = df.adm2)

# Get the names of dataframes and remove the first name
sheet_names <- names(df_list)[-1]

# Replace NA values with another value (e.g., "Missing") in each dataframe
for (sheet_name in sheet_names) {
  temp <- df_list[[sheet_name]]
  temp[is.na(temp)] <- 0
  df_list[[sheet_name]] <- temp
}


hs <- createStyle(
  textDecoration = "BOLD", fontColour = "#FFFFFF", fontSize = 12,
  fontName = "Arial Narrow", fgFill = "#95C651"
)
write.xlsx(df_list, paste0('./02_output/',output,'.xlsx'), colWidths=20,  headerStyle = hs)


# Ocha
# Write each dataframe to a separate sheet in the Excel workbook
write_xlsx(
  target.adm1 %>% select(all_of(c(adm.level[[1]],col.reached,col.totals))),
  path = paste0('./02_output/',output.ocha,'.xlsx')
)
