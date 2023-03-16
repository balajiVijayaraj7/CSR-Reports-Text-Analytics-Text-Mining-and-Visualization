#Loading necessary libraries
library(Amelia) #missing data
library(pdftools) #for reading pdf files
library(tidyverse)
library(tm) #text mining package
library(SnowballC) # for text stemming
library(tidytext)
library(textrecipes)
library(writexl) #for writing data frame to Excel sheet
library(readxl) #for reading Excel sheet


#Filtering the necessary columns from the main dataset
Energy_df <- data.frame(GRIexcel$Name, GRIexcel$Sector, GRIexcel$Size, GRIexcel$Country, GRIexcel$`Country Status`, GRIexcel$`Organization type`, GRIexcel$`Date Added`, GRIexcel$`Publication Year`, GRIexcel$`Listed/Non-listed`, GRIexcel$Region, GRIexcel$Type)

#Filtering out Energy Sector
Energy_df <- Energy_df[GRIexcel$Sector == "Energy" | GRIexcel$Sector == "Energy Utilities",]
rm(Energy_data)

#Changing Column Names
colnames(Energy_df) <- c("Name", "Sector", "Size", "Country", "Country.Status","Organization.Type", "Date.Added", "Publication.Year", "Listed.Non-Listed", "Region", "Type")

#Adding new column with Name and Publication.Year
Energy_df$filename <- paste(Energy_df$Name,Energy_df$Publication.Year, sep = "_") 

#Removes the space in between names
Energy_df$filename <- gsub(" ", "", Energy_df$filename)

#Concatinating with .pdf
Energy_df$filename <- paste(Energy_df$filename,".pdf") 
Energy_df$filename <- gsub(" ", "", Energy_df$filename) 

#Changing the type of year as factors
Energy_df$Publication.Year <- as.factor(Energy_df$Publication.Year)

#Removing duplicate files
Energy_df <- Energy_df[!duplicated(Energy_df$filename), ]

#Creating a list dataframe of filenames
Energy_df_ls_df <- as.list.data.frame(Energy_df$filename)

getwd()
setwd("/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/Energy")

#Taking all the files in the path as a list dataframe
files <- list.files("/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/Energy")
files <- as.list.data.frame(files)

# Create an empty tibble with three columns
Energy_tbl <- tibble(word = character(), count = integer(), filename = character())

# Loop through each file in the Energy_df_ls_df list and process the PDF
for (i in 1:nrow(Energy_df)) {
  if (Energy_df_ls_df[i] %in% files) {
    tryCatch({
      PDF <- pdf_text(Energy_df_ls_df[i]) %>%
        read_lines() #open the PDF inside your project folder and reads each line
      
      text_df <- tibble(line = 1:length(PDF), text = PDF) #turns each line into a tibble containing individual words
      
      text_tokens <- text_df %>%
        unnest_tokens(word, text) %>% #breaking text to individual tokens
        anti_join(stop_words) %>% #to remove stopwords
        count(word, sort = TRUE) #to count the occurrence of each word
      
      # Add the result of the loop to the tibble with the corresponding filename
      Energy_tbl <- Energy_tbl %>% 
        add_row(word = text_tokens$word, count = text_tokens$n, filename = Energy_df_ls_df[i])
    }, error = function(e) {
      # Handle the exception/error here
      cat("Error occurred: ", conditionMessage(e), "\n")
    })
  }
}

#Tibble to data frame
word_count_df <- as.data.frame(Energy_tbl) 

#dataframe to CSV
write.csv(word_count_df, "/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/word_count_df.csv", row.names=FALSE)

#csv to tibble
Energy_tbl <- read_csv("/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/word_count_df.csv")

#Removing numbers from the dataset
Energy_tbl <- Energy_tbl %>%
  filter(!grepl("[0-9]", word))

# Sum the counts for each word and grouping it based on file names
word_counts_1<- Energy_tbl %>%
  group_by(filename) %>%
  summarize(total_count = sum(count))

#Languages - English, German, French, Swedish, Spanish, Chinese, Japanese, Greek, Portuguese, Russian
#these languages contributes to 92.6% of the total available pdfs(Non-Corrupted pdfs)

#----------------Greenhouse gas Emission----------------#
#filtering key terms from the Energy_tbl
filtered_words_GE <- Energy_tbl %>% 
  filter(word == "carbon" | word == "methane"  
         | word == "greenhouse" | word == "emissions"  
         | word == "fossil" | word == "fuel" | word == "renewable"  
         | word == "ozone" | word == "energy" | word == "efficiency" 
         | word == "offset" | word == "reduction" | word == "sustainable"
         | word == "kohlenstoff" | word == "methan" | word == "distickstoffoxid"
         | word == "ozon" | word == "erneuerbar" | word == "energie"
         | word == "effizienz" | word == "ausgleich" | word == "reduktion"
         | word == "nachhaltig" | word == "treibhaus" | word == "kol"
         | word == "metan" | word == "lustgas" | word == "ozon"
         | word == "förnybar" | word == "lustgas" | word == "ozon"
         | word == "metan" | word == "energi" | word == "effektivitet"
         | word == "kompensera" | word == "minskning" | word == "hållbar"
         | word == "växthus" | word == "utsläpp" | word == "carbone"
         | word == "méthane" | word == "nitreux" | word == "ozone"
         | word == "renouvelable" | word == "efficacité" | word == "compensation"
         | word == "réduction" | word == "durable" | word == "serre"
         | word == "émission" | word == "碳" | word == "甲烷"
         | word == "氧化亚氮" | word == "臭氧" | word == "可再生"
         | word == "能源" | word == "效率" | word == "抵消"
         | word == "减少" | word == "可持续" | word == "温室"
         | word == "排放" | word == "炭素" | word == "メタン"
         | word == "亜酸化窒素" | word == "オゾン" | word == "再生可能"
         | word == "エネルギー" | word == "効率" | word == "オフセット"
         | word == "減少" | word == "持続可能" | word == "温室"
         | word == "放出" 
         | word == "invernadero" | word == "emisión" | word == "carbono"
         | word == "metano" | word == "nitroso" | word == "ozono"
         | word == "renovable" | word == "energía" | word == "ficiencia"
         | word == "compensación" | word == "reducción" | word == "sostenible"
         | word == "carbono" | word == "metano" | word == "nitroso"
         | word == "ozônio" | word == "renovável" | word == "energia"
         | word == "eficiência" | word == "compensação" | word == "redução"
         | word == "sustentável" | word == "estufa" | word == "emissões"
         | word == "θερμοκήπιο" | word == "εκπομπές" | word == "άνθρακας"
         | word == "μεθάνιο" | word == "αζωτούχο" | word == "όζον"
         | word == "ανανεώσιμος" | word == "ενέργεια" | word == "απόδοση"
         | word == "αντιστάθμιση" | word == "μείωση" | word == "βιώσιμος"
         | word == "углерод" | word == "метан" | word == "оксид"
         | word == "озон" | word == "возобновляемый" | word == "энергия"
         | word == "эффективность" | word == "компенсация" | word == "сокращение"
         | word == "устойчивый" | word == "теплица" | word == "выбросы"
  )

# Sum the counts for each word and grouping it based on file names
#word count for Greenhouse GAs Emission
word_counts_GE <- filtered_words_GE %>%
  group_by(filename) %>%
  summarize(total_count = sum(count))

#Tibble to DataFrame
word_counts_GE_df <- as.data.frame(word_counts_GE)
colnames(word_counts_GE_df) <- c("filename", "GE")

#DataFrame to csv
write.csv(word_counts_GE_df, "/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/word_counts_GE_df.csv", row.names=FALSE)

#CSV to Tibble
word_counts_GE <- read_csv("/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/word_counts_GE_df.csv")

#Outerjoin the data frame to create new dataframe
Energy_df <- merge(Energy_df, word_counts_GE_df, by = c("filename"), all=TRUE)


#-----------------Diversity-----------------------#
filtered_words_D <- Energy_tbl %>% 
  filter(word == "inclusion" | word == "equity" | word == "intersectionality" | word == "diversity"
         | word == "allyship" | word == "justice" | word == "representation" 
         | word == "underrepresented" | word == "discrimination" | word == "cultural" 
         | word == "awareness" | word == "vielfalt" | word == "inklusion" | word == "gleichheit"
         | word == "intersektionalität" | word == "unterstützung" | word == "gerechtigkeit"
         | word == "repräsentation" | word == "unterrepräsentiert" | word == "diskriminierung"
         | word == "kultur" | word == "bewusstsein" | word == "mångfald"
         | word == "inkludering" | word == "jämlikhet" | word == "intersectionalitet"
         | word == "allieradskap" | word == "rättvisa" | word == "representation"
         | word == "underrepresenterade" | word == "diskriminering" | word == "kultur"
         | word == "medvetenhet" | word == "diversité" | word == "inclusion"
         | word == "équité" | word == "intersectionnalité" | word == "aliance"
         | word == "justice" | word == "représentation" | word == "sous-représenté"
         | word == "discrimination" | word == "culturel" | word == "conscience"
         | word == "多样性" | word == "包容性" | word == "公平性"
         | word == "交叉性" | word == "盟友支持" | word == "社会正义"
         | word == "代表性" | word == "代表不足" | word == "歧视"
         | word == "文化" | word == "意识" | word == "多様性"
         | word == "包摂" | word == "公正" | word == "相差"
         | word == "味方関係" | word == "社会正義" | word == "代表性"
         | word == "非代表性" | word == "差別" | word == "文化"
         | word == "認識" 
         | word == "diversidad" | word == "inclusión" | word == "equidad"
         | word == "interseccionalidad" | word == "alianza" | word == "justicia"
         | word == "representación" | word == "subrepresentado" | word == "discriminación"
         | word == "cultural" | word == "conciencia" 
         | word == "diversidade" | word == "inclusão" | word == "equidade"
         | word == "interseccionalidade" | word == "aliança" | word == "justiça"
         | word == "representação" | word == "subrepresentado" | word == "discriminação"
         | word == "cultural" | word == "conscientização" | word == "ποικιλομορφία"
         | word == "ενσωμάτωση" | word == "ισότητα" | word == "διασταυρούμενες"
         | word == "συμμαχία" | word == "δικαιοσύνη" | word == "αντιπροσωπεία"
         | word == "υποαναπαραστατικός" | word == "διάκριση" | word == "πολιτιστικός"
         | word == "ευαισθητοποίηση" | word == "разнообразие" | word == "включение"
         | word == "равенство" | word == "интерсекциональность" | word == "союзничество"
         | word == "справедливость" | word == "представительство" | word == "дискриминация"
         | word == "культурный" | word == "осведомленность" 
  )

# Sum the counts for each word
word_counts_D <- filtered_words_D %>%
  group_by(filename) %>%
  summarize(total_count = sum(count))

#Tibble to DataFrame
word_counts_D_df <- as.data.frame(word_counts_D)
colnames(word_counts_D_df) <- c("filename", "D")

write.csv(word_counts_D_df, "/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/word_counts_D_df.csv", row.names=FALSE)
word_counts_D <- read_csv("/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/word_counts_GE_df.csv")

#Energy_df_combined <- merge(Energy_df_combined, word_counts_GE_df, by = c("filename"))
Energy_df_combined <- merge(Energy_df_combined, word_counts_D_df, by = c("filename"), all= TRUE)


#----------Employee health and Safety-----------#

filtered_words_EHS <- Energy_tbl %>% 
  filter(word == "ergonomics" | word == "wellness" | word == "employee"| word == "health"
         | word == "culture" | word == "training" | word == "safety" 
         | word == "equipment" | word == "hazard" | word == "incident" 
         | word == "prevention" | word == "mental" | word == "mitarbeiter" 
         | word == "gesundheit" | word == "ergonomie" | word == "wohlbefinden" 
         | word == "kultur" | word == "schulung"
         | word == "sicherheit" | word == "ausrüstung" | word == "gefahr"
         | word == "vorfall" | word == "prävention" | word == "psychisch"
         | word == "anställd" | word == "hälsa" | word == "ergonomi"
         | word == "välbefinnande" | word == "utbildning" | word == "säkerhet"
         | word == "utrustning" | word == "risk" | word == "incident"
         | word == "förebyggande" | word == "psykisk" | word == "employé"
         | word == "santé" | word == "être" | word == "culture"
         | word == "formation" | word == "sécurité" | word == "equipement"
         | word == "risque" | word == "incident" | word == "prévention"
         | word == "员工" | word == "健康" | word == "人机工程学"
         | word == "健康保健" | word == "文化" | word == "培训"
         | word == "安全" | word == "设备" | word == "危险"
         | word == "事故" | word == "预防" | word == "心理"
         | word == "従業員" | word == "健康" | word == "人間工学"
         | word == "ウェルネス" | word == "文化" | word == "トレーニング"
         | word == "安全" | word == "装備" | word == "危険"
         | word == "事件" | word == "予防" | word == "メンタル"
         | word == "empleado/a" | word == "salud" | word == "seguridad"
         | word == "ergonomía" | word == "bienestar" | word == "cultura"
         | word == "capacitación" | word == "seguridad" | word == "equipo"
         | word == "peligro" | word == "incidente" | word == "prevención"
         | word == "funcionário" | word == "saúde" | word == "ergonomia"
         | word == "bemestar" | word == "cultura" | word == "treinamento"
         | word == "segurança" | word == "equipamento" | word == "perigo"
         | word == "incidente" | word == "prevenção" | word == "mental"
         | word == "υπάλληλος" | word == "υγεία" | word == "εργονομία"
         | word == "ευεξία" | word == "πολιτισμός" | word == "εκπαίδευση"
         | word == "ασφάλεια" | word == "εξοπλισμός" | word == "κίνδυνος"
         | word == "περιστατικό" | word == "πρόληψη" | word == "ψυχικός"
         | word == "сотрудник" | word == "здоровье" | word == "эргономика"
         | word == "эргономика" | word == "культура" | word == "обучение"
         | word == "безопасность" | word == "оборудование" | word == "опасность"
         | word == "инцидент" | word == "предотвращение" | word == "психический"
  )

# Sum the counts for each word
word_counts_EHS <- filtered_words_EHS %>%
  group_by(filename) %>%
  summarize(total_count = sum(count))

#Tibble to DataFrame
word_counts_EHS_df <- as.data.frame(word_counts_EHS)
colnames(word_counts_EHS_df) <- c("filename", "EHS")

write.csv(word_counts_EHS_df, "/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/word_counts_EHS_df.csv", row.names=FALSE)
word_counts_EHS <- read_csv("/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/word_counts_EHS_df.csv")

#Energy_df_combined <- merge(Energy_df_combined, word_counts_GE_df, by = c("filename"))
Energy_df_combined <- merge(Energy_df_combined, word_counts_EHS_df, by = c("filename"), all= TRUE)


#----------Customer Welfare-----------# 
filtered_words_CW <- Energy_tbl %>% 
  filter(word == "ethical" 
         | word == "transparency" | word == "trade" | word == "eco"
         | word == "consumer" | word == "feedback" | word == "responsibility"
         | word == "liability" | word == "recalls"| word == "kunde"
         | word == "wohlfahrt" | word == "ethik" | word == "transparenz" 
         | word == "handel" | word == "umweltfreundlich"
         | word == "nachhaltig" | word == "verbraucher" | word == "responsabilité"
         | word == "verantwortung" | word == "haftung" | word == "rückrufe"
         | word == "kund" | word == "välfärd" | word == "etisk"
         | word == "transparens" | word == "miljövänlig" | word == "hållbar"
         | word == "konsument" | word == "ansvar" | word == "ansvarighet"
         | word == "terkallelser" | word == "client" | word == "bien-être"
         | word == "ethique" | word == "transparence" | word == "commerce"
         | word == "ecologique" | word == "durable" | word == "consommateur"
         | word == "civile" | word == "rappels" | word == "prévention"
         | word == "客户 " | word == "福利" | word == "道德"
         | word == "透明度" | word == "贸易" | word == "环保的"
         | word == "可持续的" | word == "消费者" | word == "反馈"
         | word == "责任" | word == "责任" | word == "召回"
         | word == "顧客" | word == "福祉" | word == "倫理的"
         | word == "透明性" | word == "貿易" | word == "エコフレンドリー"
         | word == "持続可能な" | word == "消費者" | word == "フィードバック"
         | word == "責任" | word == "責任" | word == "リコール"
         | word == "cliente" | word == "bienestar" | word == "ético"
         | word == "transparencia" | word == "comercio" | word == "amigable"
         | word == "sostenible" | word == "consumidor/a" | word == "retroalimentación"
         | word == "responsabilidad" | word == "責任" | word == "リコール"
         | word == "legal" | word == "retiros" 
         | word == "cliente" | word == "ético" | word == "transparência"
         | word == "comércio" | word == "ecoamigável" | word == "sustentável"
         | word == "consumidor" | word == "feedback" | word == "responsabilidade"
         | word == "civil" | word == "recalls" | word == "πελάτης"
         | word == "ηθικός" | word == "διαφάνεια" | word == "εμπόριο"
         | word == "περιβάλλον" | word == "βιώσιμος" | word == "καταναλωτής"
         | word == "ανατροφοδότηση" | word == "ευθύνη" | word == "αμέλεια"
         | word == "ανακλήσεις" | word == "клиент" | word == "этичный"
         | word == "прозрачность" | word == "торговля" | word == "экологичный"
         | word == "устойчивый" | word == "потребитель" | word == "обратная"
         | word == "ответственность" | word == "юридическая" | word == "отзывы"
  )

# Sum the counts for each word
word_counts_CW <- filtered_words_CW %>%
  group_by(filename) %>%
  summarize(total_count = sum(count))

#Tibble to DataFrame
word_counts_CW_df <- as.data.frame(word_counts_CW)
colnames(word_counts_CW_df) <- c("filename", "CW")

write.csv(word_counts_CW_df, "/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/word_counts_CW_df.csv", row.names=FALSE)
word_counts_CW <- read_csv("/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/word_counts_CW_df.csv")

#Energy_df_combined <- merge(Energy_df_combined, word_counts_GE_df, by = c("filename"))
Energy_df_combined <- merge(Energy_df_combined, word_counts_CW_df, by = c("filename"), all= TRUE)


Energy_df_combined <- Energy_df_combined[,-15]   

#Exporting the dataframe to Excel
write_xlsx(Energy_df_combined,"/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/Energy_df_combined.xlsx")


#importing from excel
Energy_df_combined <- read_excel("/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/Energy_df_combined.xlsx", sheet = "Sheet1")


#-------------------Data For WordCloud--------------------#
# Sum the counts for each word
word_cloud_GE <- filtered_words_GE %>%
  group_by(word) %>%
  summarize(total_count = sum(count))

word_cloud_GE_df <- as.data.frame(word_cloud_GE)
colnames(word_cloud_GE_df) <- c("GE_Word", "Total_count")

# Sum the counts for each word
word_cloud_D <- filtered_words_D %>%
  group_by(word) %>%
  summarize(total_count = sum(count))

word_cloud_D_df <- as.data.frame(word_cloud_D)
colnames(word_cloud_D_df) <- c("D_Word", "Total_count")

# Sum the counts for each word
word_cloud_EHS <- filtered_words_EHS %>%
  group_by(word) %>%
  summarize(total_count = sum(count))

word_cloud_EHS_df <- as.data.frame(word_cloud_EHS)
colnames(word_cloud_EHS_df) <- c("ord", "Total_count")

# Sum the counts for each word
word_cloud_CW <- filtered_words_CW %>%
  group_by(word) %>%
  summarize(total_count = sum(count))

word_cloud_CW_df <- as.data.frame(word_cloud_CW)
colnames(word_cloud_CW_df) <- c("CW_Word", "Total_count")

#Creating DataFrame for WordCloud
max_rows <- max(nrow(word_cloud_GE_df), nrow(word_cloud_D_df), nrow(word_cloud_EHS), nrow(word_cloud_CW))

word_cloud_D_df[((nrow(word_cloud_D_df) + 1):max_rows), ] <- NA
word_cloud_EHS[((nrow(word_cloud_EHS) + 1):max_rows), ] <- NA
word_cloud_CW[((nrow(word_cloud_CW) + 1):max_rows), ] <- NA

Word_count_data_frame <- cbind(word_cloud_GE_df, word_cloud_D_df, word_cloud_EHS, word_cloud_CW)
colnames(Word_count_data_frame) <- c("GE", "GE_Total_count", "D", "D_Total_count", "EHS", "EHS_Total_count", "CW", "CW_Total_count")

#Exporting word count dataframe to excel
write_xlsx(Word_count_data_frame,"/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/Word_count_data_frame.xlsx")
#-----------------------------------------------------------#

Energy_df_combined$Total <- rowSums(Energy_df_combined[, c("GE", "D", "EHS", "CW")], na.rm = TRUE)
Energy_df_combined <- Energy_df_combined[,-17]

#removing duplicate files
Energy_df_combined <- Energy_df_combined[!duplicated(Energy_df_combined$filename), ]

#exporting finalised data frame to excel
write_xlsx(Energy_df_combined,"/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/Energy_df_combined.xlsx")

#-----------------------------END---------------------#