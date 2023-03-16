install.packages("scales")
install.packages("ggplot2")
library(ggplot2)
library(scales)
library(ggmap)


#Importing the dataset
Energy_df_combined <- read_excel("/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/Energy_df_combined.xlsx", sheet = "Sheet1")
Energy_df_combined <- Energy_df_combined[!duplicated(Energy_df_combined$filename), ]

#Converting the columns to factors
Energy_df_combined$Size <- as.factor(Energy_df_combined$Size)
Energy_df_combined$Country <- as.factor(Energy_df_combined$Country)
Energy_df_combined$Country.Status <- as.factor(Energy_df_combined$Country.Status)
Energy_df_combined$Organization.Type <- as.factor(Energy_df_combined$Organization.Type)
Energy_df_combined$Publication.Year <- as.factor(Energy_df_combined$Publication.Year)
Energy_df_combined$`Listed.Non-Listed` <- as.factor(Energy_df_combined$`Listed.Non-Listed`)
Energy_df_combined$Region <- as.factor(Energy_df_combined$Region)
Energy_df_combined$Type <- as.factor(Energy_df_combined$Type)
Energy_df_combined$Sector <- as.factor(Energy_df_combined$Sector)

Energy <- Energy_df_combined

Energy <- Energy[!duplicated(Energy$filename), ]
summary(Energy)

#---------------------PLOT-1--------------------------------#
#---------------Attention level by countries--------------#
#creating data subset for plotting
data <- Energy%>%
  select(Country, D, GE, EHS, CW, Total)%>%
  group_by(Country)%>%
  summarise(
    Greenhouse = sum(GE,na.rm = TRUE),
    Diversity = sum(D, na.rm = TRUE), 
    Employee = sum(EHS, na.rm = TRUE), 
    Customer = sum(CW, na.rm = TRUE),
    Total = sum(Total, na.rm = TRUE))

#Scaling the Total column
new_total <- data$Total
new_total <- rescale(new_total, to=c(100,1000))
new_total <-as.data.frame(new_total)

data <- cbind(data, new_total)

#Getting the map data 
mapdata<- map_data("world")

colnames(data)[colnames(data)=="Country"]<-"region"

data$region <- as.character(data$region)

#Modifying the country names to match the names in map data
idx <- which(data$region == "Korea, Republic of")
data[idx, "region"] <- "South Korea"

idx_1 <- which(data$region == "Mainland China")
data[idx_1, "region"] <- "China"

idx_2 <- which(data$region == "Russian Federation")
data[idx_2, "region"] <- "Russia"

idx_3 <- which(data$region == "Slovak Republic")
data[idx_3, "region"] <- "Slovakia"

idx_4 <- which(data$region == "United States of America")
data[idx_4, "region"] <- "USA"

idx_5 <- which(data$region == "Viet Nam")
data[idx_5, "region"] <- "Vietnam"

idx_6 <- which(data$region == "United Kingdom of Great Britain and Northern Ireland")
data[idx_6, "region"] <- "UK"

data$region <- as.factor(data$region)

#Combining map data and data subset
mapdata_1 <- left_join(mapdata,data, by = "region")
mapdata_1[is.na(mapdata_1)] <- 0

#exporting data frame to excel
write_xlsx(mapdata_1,"/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/mapdata_1.xlsx")

mapdata_1 <- read_excel("/Users/balajivijayaraj/Desktop/DAV/Home_Assignment/mapdata_1.xlsx", sheet = "Sheet1")


# Adding a title and subtitle to the plot
title <- "Level of Attention Given to Address Key Issues in CSR by different Countries"
subtitle <- "Based on Greenhouse Gas Emission, Employee Health and Safety, Customer Welfare and Diversity"

# Adjusting the color scale
col <- c("#EFF3FF", "#BDD7E7", "#6BAED6", "#3182BD", "#08519C")
scale_fill <- scale_fill_gradientn(name = "Level of Attention",
                                   colors = col, na.value = "white")

# Adjusting the legend
legend <- theme(legend.title = element_text(size = 13, face = "bold", margin = margin(b = 10)),
                legend.position = "right",
                legend.text = element_text(size = 10, face = "bold"))

#Annotation Label
continent_labels <- data.frame(
  continent = c("Africa", "Asia", "Europe", "North America", "South America", "Australia"),
  lon = c(24, 100, 18, -100, -56, 135),
  lat = c(0, 35, 50, 40, -14, -25)
)

#Theme for the plot
theme_custom <- theme_bw() +
  theme(panel.border = element_blank(),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank(),
        axis.line = element_blank(),
        axis.text = element_blank(),
        axis.ticks = element_blank(),
        axis.title.y = element_blank(),
        axis.title.x = element_blank(),
        plot.title = element_text(size = 18, face = "bold", hjust = 0.5),
        plot.subtitle = element_text(size = 14, hjust = 0.5))

# Creating the plot
world_map <- ggplot(mapdata_1, aes(x = long, y = lat, group = group)) +
  geom_polygon(aes(fill = new_total), color = "black") +
  scale_fill +
  labs(title = title, subtitle = subtitle) + 
  
  annotate(geom = "text", data = continent_labels, x = continent_labels$lon, y = continent_labels$lat,
           label = continent_labels$continent, size = 4, fontface = "bold", color = "black") +
  theme_custom +
  legend

# Displaying the plot
world_map

#------------------------------PLOT-2-----------------------#
#---------------Attention level to different areas by types of company-------#

data_2 <- Energy %>% 
  select(Type, D, GE, EHS, CW, Total) %>% 
  group_by(Type) %>% 
  summarise(
    GE = sum(GE, na.rm = TRUE),
    D = sum(D, na.rm = TRUE), 
    EHS = sum(EHS, na.rm = TRUE), 
    CW = sum(CW, na.rm = TRUE),
    Total = sum(Total, na.rm = TRUE))


# reshape data for plotting
data_3 <- data_2 %>% 
  select(Type, GE:CW) %>% 
  tidyr::gather(key = "Category", value = "Total", -Type)

# create the plot
ggplot(na.omit(data_3), aes(x = Type, y = Total, fill = Category)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.85), color = NA, width = 0.8) +
  scale_fill_manual(values = c("#00AFBB", "#FC4E07", "#F2A900", "#7FBC41"),
                    name = "Category") +
  labs(title = "Level of attention given to different areas by different types of Companies",
       x = "Types of Companies", y = "Frequency of Key Issues Mentioned (in Thousands)") +
  theme_minimal() +
  scale_y_continuous(limits = c(0, 110000), breaks = seq(0, 110000, by = 10000), labels = seq(0, 110, by = 10)) +
  theme(axis.text.x = element_text(hjust = 0.45, size = 12, vjust = 5, face = "bold"),
        axis.text.y = element_text(size = 12),
        axis.title = element_text(size = 15, face = "bold"),
        panel.grid.major.y = element_line(color = "gray90", linetype = "dashed"),
        panel.grid.major.x = element_blank(),
        panel.grid.minor = element_blank(),
        legend.title = element_text(size = 12, face = "bold", margin = margin(b = 5)),
        legend.text = element_text(size = 11),
        legend.key.size = unit(0.6, "cm"),
        plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
        plot.margin = unit(c(1, 1, 1, 0.5), "cm"))


#----------------------PLOT-3---------------------#
#--------------Attention given to different areas over different points of time------------------#
data_4 <- (Energy) %>% 
  select(Publication.Year, D, GE, EHS, CW, Total) %>% 
  group_by(Publication.Year) %>% 
  summarise(
    GE = sum(GE, na.rm = TRUE),
    D = sum(D, na.rm = TRUE), 
    EHS = sum(EHS, na.rm = TRUE), 
    CW = sum(CW, na.rm = TRUE),
    Total = sum(Total, na.rm = TRUE))

#Execute it twice to remove unwanted rows(2018 and NA)
data_4 <- data_4[-16,]

#creating the plot
data_4 %>%
  ggplot(aes(x = Publication.Year)) +
  geom_line(aes(y = D, color = "Diversity", group = 1), linewidth = 1.25) +
  geom_line(aes(y = GE, color = "Greenhouse Gas Emission", group = 1), linewidth = 1.25) +
  geom_line(aes(y = EHS, color = "Employee Health and Safety", group = 1), linewidth = 1.25) +
  geom_line(aes(y = CW, color = "Customer Welfare", group = 1), linewidth = 1.25) +
  geom_line(aes(y = Total, color = "Total", group = 1), linewidth = 1.5) +
  scale_color_manual(name = "Key Issues", values = c("Diversity" = "#FF5733", "Greenhouse Gas Emission" = "#00BFFF", "Employee Health and Safety" = "#00FF7F", "Customer Welfare" = "#FFC300", "Total" = "#800080")) +
  xlab("Year") +
  ylab("Word Frequency (in Thousands)") +
  ggtitle("Level of attention given to different areas over different points of time") +
  theme_minimal() +
  scale_y_continuous(limits = c(0, 150000), breaks = seq(0, 150000, by = 10000), labels = seq(0, 150, by = 10)) +
  theme(
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank(),
    panel.background = element_blank(),
    axis.line = element_line(colour = "black", size = 0.5),
    axis.text = element_text(size = 12),
    axis.title = element_text(size = 14, face = "bold"),
    legend.title = element_text(size = 13, face = "bold", margin = margin(b = 5)),
    legend.text = element_text(size = 12),
    plot.title = element_text(size = 16, face = "bold", hjust = 0.5)
  )


#-----------------------PLOT-4--------------------------#
#-----------Attention given to different areas by region and size of company-------#
data_5 <- Energy %>% 
  select(Region,Size, D, GE, EHS, CW, Total)

#creating the plot
ggplot(na.omit(data_5), aes(x = Region, y = Total, fill = Size)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.9), color = NA, width = 0.8) +
  scale_fill_manual(values = c("#00AFBB", "#FC4E07", "#F2A900", "#7FBC41"),
                    name = "Size") +
  labs(title = "Level of attention given to different areas by region and size of company",
       x = "Region", y = "Frequency of Key Issues Mentioned") +
  facet_wrap(~Size, scales = "free_y", ncol = 1) +
  theme_minimal() +
  theme(axis.text.x = element_text(hjust = 0.45, size = 11, vjust = 3, face = "bold",margin = margin(b = 5)),
        axis.text.y = element_text(size = 11),
        axis.title = element_text(size = 14, face = "bold"),
        panel.grid.major.y = element_line(color = "gray90", linetype = "dashed"),
        panel.grid.major.x = element_blank(),
        panel.grid.minor = element_blank(),
        legend.title = element_text(size = 12, face = "bold"),
        legend.text = element_text(size = 11 ),
        legend.key.size = unit(0.6, "cm"),
        plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
        plot.margin = unit(c(1, 1, 1, 0.5), "cm"),
        panel.spacing = unit(1, "lines"),
        strip.text = element_text(size = 12, face = "bold")) +
  guides(fill = FALSE)

#--------------------PLOT-5-------------------------------#
#-----Attention given to Key Issues by Region and Country Status------#

data_7 <- Energy %>% 
  select(Region,Country.Status, D, GE, EHS, CW, Total)

# create plot using ggplot2
ggplot(na.omit(data_7), aes(x = Region, y = Total, fill = Country.Status)) +
  geom_col(position = "stack") +
  scale_fill_manual(values = c("#FFC300", "#C70039", "#636EFA", "#8B5B3E", "#00BFFF")) +
  labs(x = "Region", y = "Attention Given to Key Issues (in Thousands)", fill = "Country Status", title = "Attention given to Key Issues by Region and Country Status") +
  scale_y_continuous(limits = c(0, 300000), breaks = seq(0, 300000, by = 10000), labels = seq(0, 300, by = 10)) +
  theme_bw() +
  theme(plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
        axis.text = element_text(size = 10, face = "bold"),
        axis.text.x = element_text(size = 11, face = "bold", vjust = 1.5),
        axis.title = element_text(size = 12, face = "bold"),
        legend.title = element_text(size = 12, face = "bold",  margin = margin(b = 5)),
        legend.text = element_text(size = 10),
        legend.position = c(0.80,0.85),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank(),
        panel.border = element_blank(),
        panel.background = element_blank()) +
  guides(fill = guide_legend(title = "Country Status", 
                             label.position = "bottom",
                             label.hjust = 0.5, 
                             title.position = "top", 
                             title.hjust = 0.5,
                             nrow = 1)) +
  theme(
    legend.text = element_text(size = 10),
    legend.key.size = unit(0.8, "cm"),
    legend.key.height = unit(0.4, "cm"))

 #-----------------------------------END---------------------------------------#



