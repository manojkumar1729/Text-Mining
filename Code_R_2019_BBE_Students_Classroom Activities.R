#We used the same code for all our work to get bi-grams & tri-grams by changing Excel data paths while choosing 
library(tidyverse)
library(readxl)
library(ggplot2)
library(lattice)
library(caret)
library(dplyr)
library(stopwords)
library(tm)
library(tidytext)
library(hunspell)
library (corpus)
library(openxlsx)

#choosing the Excel data from PC
path <- "/Users/Manoj/Desktop/Drexel_Academics/MSBA/MIS 612/Text Mining/Final Presentation/Excel sheets/student_classroomactivities.xlsx"

# getting data from sheets
sheets <- openxlsx::getSheetNames(path)
data_frame <- lapply(sheets, openxlsx::read.xlsx, xlsxFile=path)

# assigning names to data frame
names(data_frame) <- sheets

# printing the data
print (data_frame)
CA_data_2019_BBE<- data_frame$"2019 BBE"
CA_data_2019_BBE
my_data = CA_data_2019_BBE

# Remove Stop words, Numbers & Symbols
text_df <- my_data$`How.did.your.classroom.activities.prepare.you.for.co-op?.If.they.didn’t,.how.were.you.prepared.for.co-op?`
text_df
text_df <- removeWords(text_df, stopwords())
text_df <- removeWords(text_df, stopwords())
text_df = gsub("[^[:alnum:]]", " ", text_df)
text_df = gsub('[[:digit:]]+', '', text_df)
text_df

# Remove extra spaces that formed after removing stop words, numbers & symbols
searchString <- '   '
replacementString <- ' '
sentenceString = gsub(searchString,replacementString,text_df)
sentenceString

searchString <- '  '
replacementString <- ' '
sentenceString = gsub(searchString,replacementString,sentenceString)
sentenceString

searchString <- '  '
replacementString <- ' '
sentenceString = gsub(searchString,replacementString,sentenceString)
sentenceString

# Refining data again depending on frequency data - remove single alphabets like a, n, x, i, t etc
sentenceString <- gsub('N ','',sentenceString)
sentenceString #Verify at every step whether the above line helped us to achieve our goal to remove "N"
sentenceString <- gsub('A','',sentenceString)
sentenceString
sentenceString <- gsub('I ','',sentenceString)
sentenceString
sentenceString <- gsub(' s ',' ',sentenceString)
sentenceString
sentenceString <- gsub('xD','',sentenceString)
sentenceString
sentenceString <- gsub(' t ',' ',sentenceString)
sentenceString

#adding new column with our refined data for further data processing
my_data$new <- sentenceString
my_data

#Finding Frequency
my_data %>%
  select(`new`) %>%
  unnest_tokens(word, `new`) %>%
  count(word, sort = TRUE)

#stemming
stemming <- my_data %>%
  unnest_tokens(word, `new`) %>%
  mutate(word = corpus::text_tokens(word, stemmer = "en") %>% unlist()) %>% # add stemming process
  count(word) %>% 
  group_by(word) %>%
  summarize(n = sum(n)) %>%
  arrange(desc(n))
stemming
write.csv(stemming,"2019_BBE_stemming.csv")       #Extracting data into new csv file

#Bi-grams
bigram_list <- my_data %>%
  unnest_tokens(bigram, `new`, token = "ngrams", n = 2) %>%  
  separate(bigram, c("word1", "word2"), sep = " ") %>%               
  unite("bigram", c(word1, word2), sep = " ") %>%
  count(bigram) %>%
  filter(n >= 5) %>%     #filter for bi-grams from words repeated 5 or more times
  pull(bigram)
bigram_list
write.csv(bigram_list,"2019_BBE_bigrams.csv")     #Extracting data into new csv file

#Tri-grams
trigram_list <- my_data %>%
  unnest_tokens(trigram, `new`, token = "ngrams", n = 3) %>%  
  separate(trigram, c("word1", "word2", "word3"), sep = " ") %>%               
  unite("trigram", c(word1, word2, word3), sep = " ") %>%
  count(trigram) %>%
  filter(n >= 3) %>%    #filter for tri-grams from words repeated 3 or more times
  pull(trigram)
trigram_list
write.csv(trigram_list,"2019_BBE_trigrams.csv")    #Extracting data into new csv file
