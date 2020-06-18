library(tidyverse)
library(readxl)
library(lubridate)
library(janitor)

#Import dataset
Qualtrics <- read_excel(list.files(pattern="*.xlsx"), skip = 1)

#Select and rename columns
quantitative_data_columns <- c('Recorded Date', 'Agreement number', 'Agency Name', 'Reporting Period (Current)',
                               'Number of participants referred to the program during the current reporting period',
                               'Number of participants screened for program eligibility during the current reporting period',
                               'Number of participants served in the ARI program, during the current reporting period - Total number of participants being served during this reporting period',
                               'Number of participants served in the ARI program, during the current reporting period - Number of new enrollments during this reporting period',
                               'Number of exits that completed conditions during the current reporting period',
                               'Number of exits that were revoked during the current reporting period - Revoked to jail',
                               'Number of exits that were revoked during the current reporting period - Revoked to IDOC',
                               'Number of exits that were revoked during the current reporting period - Revoked to neither IDOC or jail',
                               'Number of exits due to other reasons (e.g., transfer, death, illness, other) during the current reporting period',
                               'Participant capacity of your program',
                               'Number of participants receiving the following types of services as part of their service plan - Participants - Cognitive behavioral therapy (CBT) - Number',
                               'Number of participants receiving the following types of services as part of their service plan - Participants - Substance use disorder (SUD) treatment - Number',
                               'Number of participants receiving the following types of services as part of their service plan - Participants - Mental health treatment - Number',
                               'Number of participants receiving the following types of services as part of their service plan - Participants - Education/employment training and supports - Number',
                               'Number of ARI-funded positions - Numer of positions filled',
                               'Number of ARI-funded positions - Number of vacant positions'
                               )

quantitative_data_columns_renamed <- c('Recorded Date', 'Agreement number', 'Agency Name', 'Reporting Period',
                                       'Referral', 'Screened', 'Total_served', 'New_enrollments',
                                       'Exits_completed', 'Exits_fail_jail', 'Exits_fail_IDOC', 'Exits_fail_other', 'Exits_other',
                                       'Program_capacity',
                                       'CBT', 'SUD', 'MH', 'Education_training',
                                       'ARI_funded_filled',
                                       'ARI_funded_vacant'
                                       )

#Create subset
quantitative_data <- Qualtrics %>% 
  filter(`Contact - Name of person submitting form` != 'Mary Ann Dyar') %>% 
  select(quantitative_data_columns)

colnames(quantitative_data) <- quantitative_data_columns_renamed

quantitative_data <- quantitative_data %>% 
  filter(`Reporting Period` == 'SFY20 Q3') %>% 
  mutate(Exits_total = `Exits_completed` + `Exits_fail_jail` + `Exits_fail_IDOC` + `Exits_fail_other` + `Exits_other`,
         Carrier_over = `Total_served` - `New_enrollments`) %>% 
  mutate(`Agreement number` = as.character(`Agreement number`), 
         CBT = as.numeric(CBT),
         SUD = as.numeric(SUD),
         MH = as.numeric(MH),
         Education_training = as.numeric(Education_training)) %>% 
  arrange(desc(`Recorded Date`)) %>% 
  distinct(`Agreement number`, `Agency Name`, `Reporting Period`, .keep_all = TRUE) %>% 
  adorn_totals('row', fill = "-", na.rm = TRUE,
               name = "Total")

#Write to csv
write.csv(quantitative_data, file = 'quantitative_data.csv')