library(readxl)
library(dplyr)
library(tidyverse)
library(data.table)
# library(openxlsx)
# library(writexl)
Sys.setenv(JAVA_HOME="**location where JAVA is stored**")
library(xlsx)

# options(encoding="utf-8")

# Question in column P:
# "If yes to Column L, % in other location"
# WAS NOT ASKED IN FRENCH

# Impossible to read from the network


setwd("**set your working directory**")
options(scipen=999)

file.list <- list.files(pattern='*.xlsx') # recursive = TRUE: to analyze subfolders
df.list1 <- lapply(file.list, read_excel)
# df.list1 <- lapply(file.list, xlsx::read.xlsx(encoding="UTF-8", stringsAsFactors = F))
df1 <- bind_rows(df.list1)


# First I will extract the time
t = file.info(file.list)
t <- t %>%
  select("ctime")%>%
  mutate(id = paste("file", 1:length(file.list), sep="" ))



test <- df1 %>%
  filter(grepl('Company Name:', `Supply and Demand Projections for Critical Drugs`) | 
           grepl('Contact', `...7`))%>% 
  select_if(~sum(!is.na(.)) > 0)%>%
  rename(col1 = `Supply and Demand Projections for Critical Drugs`,
         col2 = `...3`,                                            
         col3 = `...7`,                                            
         col4 =`...9`)%>%
  mutate(id = paste("file", rep(1:length(file.list), each=3), sep="" ))

unique(test$...3)
unique(df1$...3)
companies <- data.frame(unique(test$col2))

verificationfile <- test %>%
  select(col2, id)%>%
  filter(!is.na(col2))

write.csv(verificationfile, "companies.csv")

# Creating 2 dataframes that I will rowbind 
new1 <- test %>%
  filter(grepl('Company Name:', col1)) %>%
  select(col1, col2, id)

new <- test %>%
  select(col3:id) %>%
  rename(col1=col3, col2=col4)


d <- new %>%
  bind_rows(new1)%>%
  arrange(id)


mytype <- c(
  "text",
  "text",
  "text",
  "text",
  "text",
  "text",
  "text",
  "numeric",
  "numeric",
  "numeric",
  "numeric",
  "text", # "logical",
  "numeric",
  "numeric",
  "numeric",
  "text", # "numeric",
  "text", # "logical",
  "numeric",
  "text",
  "text")


# I need to work on this loop as it is looping twice
df.list <- lapply(file.list,function(x) {
  # creating a new column
  newcol <- paste("file", 1:length(file.list), sep="")
  dfs <- lapply(file.list, function(y) {
    read_excel(y, skip=14, range = "A16:T150", col_types = mytype) #skip=14, 
  })
  names(dfs) <- newcol
  dfs
})


df <- rbindlist(lapply(df.list, rbindlist, id = "id"))


df <- distinct(df)
df <- df %>%
  filter(!grepl('Eg', DIN) & !grepl('Add more rows as required', DIN))

fin1 <- merge(x = d, y = df, by = "id", all = TRUE) 
fin2 <- merge(x = t, y = fin1, by ="id", all=TRUE)

final <- spread(fin2, col1, col2)

finalx <- final %>% 
  filter(!is.na(DIN), 
         !is.na(`Molecule name`))


# Safety check
# f<-data.frame(unique(final$`Company Name:`), stringsAsFactors = FALSE)
# f1<-data.frame(unique(finalx$`Company Name:`), stringsAsFactors = FALSE)
# f3=f$unique.final..Company.Name...[!(f$unique.final..Company.Name... %in% f1$unique.finalx..Company.Name...)]

# rm(d, df, df.list, fin1)
#####

#CLEANING finalx

for (i in colnames(finalx)){
  finalx[[i]]=tolower(finalx[[i]])
  }

finalx <- data.frame(lapply(finalx, function(x) {
                    gsub("n/a", NA, x)
              }))


finalxx <- data.frame(lapply(finalx, as.character), stringsAsFactors=FALSE)
# finalx$Molecule.name = as.character(finalx$Molecule.name)


unique(finalxx[[19]])
finalxx[[19]]=na_if(finalxx[[19]], "n - bulk tablet manufacturer is confirming supply")
finalxx[[19]]=na_if(finalxx[[19]], "-")
finalxx[[19]]=na_if(finalxx[[19]], "unknown")
finalxx[[19]]=na_if(finalxx[[19]], "tbd")
unique(finalxx[[19]])

unique(finalxx[[14]])
finalxx[[14]] = na_if(finalxx[[14]], "y and n")
finalxx[[14]] <- gsub("y-see comments", "y", finalxx[[14]])
finalxx[[14]] <- gsub("y (distributor)", "y", finalxx[[14]])
finalxx[[14]] <- gsub("no", "n", finalxx[[14]])
finalxx[[14]] <- gsub("\\s*\\([^\\)]+\\)", "y", finalxx[[14]])
finalxx[[14]] <- gsub("ny", "n", finalxx[[14]])
finalxx[[14]] <- gsub("yy", "y", finalxx[[14]])
finalxx[[14]] <- gsub("nne", "y", finalxx[[14]])
unique(finalxx[[14]])

unique(finalxx[[18]])
finalxx[[18]] = na_if(finalxx[[18]], "na")
finalxx[[18]] = na_if(finalxx[[18]], "unknown")
finalxx[[18]] = na_if(finalxx[[18]], "-")
unique(finalxx[[18]])

for (i in colnames(finalxx)){
  finalxx[[i]]=tolower(finalxx[[i]])
}

comp <- data.frame(unique(finalxx$Company.Name.))

finalxx %>% filter(DIN=="'02237224")

finalxx <- finalxx %>% 
  mutate(DIN = replace(DIN, DIN == "'02237224", "02237224"))

# Adding leading zeros
finalxx$DIN <- sprintf("%08s", finalxx$DIN)

write.xlsx(finalxx, "final_file_55.xlsx")
write_rds(finalxx, "final_file_55.rds")


