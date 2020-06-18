library(tidyverse)
library(readxl)
library(lubridate)

#Import dataset
Qualtrics <- read_excel(list.files(pattern="*.xlsx"), skip = 1)

#Select and rename columns
qualitative_data_columns <- c('Recorded Date', 'Agreement number', 'Agency Name', 'Reporting Period (Current)',
                              'Describe the mix of sanctions and incentives used by the program',
                              'Describe any specific ways the community was involved in the program this reporting period',
                              'Briefly describe how the knowledge will be used to benefit the ARI program.',
                              'Note any unmet needs:',
                              'Briefly describe progress this reporting period toward the goals and objectives.',
                              'If "No", please explain....147',
                              'If "No", please explain....149',
                              'What problems/barriers did you encounter during the reporting that affected your ability to reach the goals and objectives listed in Exhibit F?',
                              'If "Yes", what has been your experience? Are additional resources needed to serve this expanded population?',
                              'What major activities were accomplished this reporting period?',
                              'What major activities are planned for the next reporting period?',
                              'Is there additional support that we can provide in the form of resources or technical assistance?',
                              'Please provide one participant impact or success story (optional).'
                              )

qualitative_data_columns_renamed <- c('Recorded Date', 'Agreement number', 'Agency Name', 'Reporting Period',
                                      'Sanctions and Incentives',
                                      'Community Involvement',
                                      'Knowledge Utilization',
                                      'Unmet Needs',
                                      'Goals and Objectives',
                                      'Budget Expansion',
                                      'Reduction Goal',
                                      'Barriers',
                                      'Violence Eligibility',
                                      'Major Activities_current',
                                      'Major Activities_next',
                                      'Addition Support',
                                      'Success Stories'
                                      )

#Create subset
qualitative_data <- Qualtrics %>% 
  filter(`Contact - Name of person submitting form` != 'Mary Ann Dyar') %>% 
  select(qualitative_data_columns)

colnames(qualitative_data) <- qualitative_data_columns_renamed

qualitative_data <- qualitative_data %>% 
  arrange(desc(`Recorded Date`)) %>% 
  distinct(`Agreement number`, `Agency Name`, `Reporting Period`, .keep_all = TRUE)

#Write to csv
write.csv(qualitative_data, file = 'qualitative_data.csv')
