# VBA_Misc_Macros
 This repo will contain various small macros and use cases i've implemented for my current work in order to speed up the process

# Simple explanations

- extractDt will format a 6digit number into a date format agreed upon
- Beautify will concatenate the second colummn based on if exista data in the second column, while cleaning the file in an efficient manner
- Translate() is a function that translates text input into a cell into another language, for international comunications and for the sake of keeping the openend apps at minimum. 
- GetPayments() will fetch data from a master file and will copy it into another file, based on a speciffic row and column, while keeping everything tight and clean. 

# Observations

- Beautify shaved off around 20-40 minutes (60 minutes in the most extreme case) of manual labor trying to correct a concatenate excel formula. 
- Translate did not improve the time, but the efficiency of the agent in international communications and kept the processes at minimum, mostly because the hardware allocated was slow and did not allow more apps or Chrome tabs opened at the same time. 
- GetPayments freed around 30 minuted per document where the data would be entered manually. The whole process consisted of creating and populating 455 documents manually. This single operation took approximatelly 300 hours of labor that had been saved and redistributed to other priorities, as well as other agents freed of this task
