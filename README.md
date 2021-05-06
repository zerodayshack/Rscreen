#This is the README file to run the R script Screener.R for the evaluation of the SEIS candidate
#License: Creative commons. Please cite this scritp with Barbara Russo, Free University of Bozen-Bolzano 2021 "Screener.R" 

Prerequisite

Install R or RStudio
Create a folder in which to save Screener.R, say "EvaluationApp"
Save the full path name to the folder EvaluationApp
Import the folder (e.g., called "ExcelFiles") with the candidates' excel files in the folder EvaluationApp
Import the file with the university rankings in the folder EvaluationApp
Import the file with the list of candidates (from the Secretariat) in the folder EvaluationApp

Run from terminal:

Rscript Screener.R <path to folder EvaluationApp> <name of the folder ("ExcelFiles")> <name of the file candidates list ("Lista_candidati_SEIS_1sessione_")> <name and suffix of the university ranking file ("JANUARY2020.xlsx")>

If you do not want to run it from terminal uncomment the corresponding lines in the script. 

Output:
Screener.R creates a folder named "Evaluation" in which there are two files a csv file and a txt file, for example

File1= _2020-05-19_evaluation.csv

File 2= itemsToReview.txt

The name of File 1 is generated based on the date of the system followed by the term"evaluation.csv":  "_Date_evalution.csv"
The name of File 2 is always "itemsToReview.txt"

Everytime Screener.R is run, the two files are overridden. The timestamp of the evaluation is reported at the bottom line of each file. 

File 1 contains the evaluation of each candidate with 
NAME	SURNAME	BEING_NEUROPEAN	DURATION_BSC	WEIGHTED_GRADE	TYPE_BSC	WORKING_EXPERIENCE_MONTHS	TOTAL_SCORE	RESULT	STILL_TO_GRADUATE

File 2 contains warnings for each candidate. The system warns the commission on item that need to be checked manually or are not clear.

Screener.R 
automatically converts any grading system (either final grade or through conversion of individual courses) to CGPA (4-1) 
evaluates weather the working experience is in the field of interest through an embedded vocabulary.
checks whether the candidate excel file corresponds to a name in the list of candidates also swapping names with surname or splitting the names. 
checks whether the length of the BSc is adequate and the BSc is eligible again using an internal vocabulary.  
and verifies whether the candidate is admitted by checking the total score greater than 5.

Variables are set to the following default values

weights for the total score
wStud<-0.6
wBSc<-0.3
wWE<-0.1
wCV<-1
#This is a default value for the CV. It will not be changed by the R script. 
CV<-5

#Type of BSc degree. Default 0
#number of months of work. Default 0
#Still to graduate? Default False
#Duration of the BSc. Default empty char
#degree grade in CGPA. Default is 0
#final score of the candidate. Default is 5. 6 to pass



