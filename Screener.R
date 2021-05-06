#Check installed packages
load_or_install_package <- function(){
  if("stringr" %in% rownames(installed.packages()) == FALSE){install.packages("stringr") + library(stringr)}else{
    library(stringr)
    if( "gridExtra" %in% rownames(installed.packages()) == FALSE ){ install.packages("gridExtra") + library(gridExtra)}else{
      library(gridExtra)
      if("grid" %in% rownames(installed.packages()) == FALSE){install.packages("grid") + library(grid)}else{
        library(grid)
        if("here" %in% rownames(installed.packages()) == FALSE){install.packages("here")  +   library(here)}else{
          library(here)
          if("Rmpfr" %in% rownames(installed.packages()) == FALSE){install.packages("Rmpfr") + library(Rmpfr)}else{
            library(Rmpfr)
            if("xtable" %in% rownames(installed.packages()) == FALSE){install.packages("xtable") + library(xtable)}else{
              library(Rmpfr)
              if("nls2" %in% rownames(installed.packages()) == FALSE){install.packages("nls2") + library(nls2)}else{
                library(nls2)
                if("readxl" %in% rownames(installed.packages()) == FALSE){install.packages("readxl") + library(readxl)}else{
                  library(readxl)
                  if("optparse" %in% rownames(installed.packages()) == FALSE){install.packages("optparse") + library(optparse)}else{
                    library(optparse)
                }
              }}}}}}}}
}
load_or_install_package()


#OPTIONS two decimals and no scientific notaion in numbers
options(digits=6,scipen=999)
defaultW <- getOption("warn")
options(warn = -1)

#Clean up variables
#rm(list = ls())
setwd("~/")
#uncomment these lines to use it in command line
# args = commandArgs(trailingOnly=TRUE)
# candidatesFolder<-args[1]
# setwd(candidatesFolder)
# excelFilesFolderPattern<-args[2]
# listCandsPattern<-args[3]
# uRankingFile<-args[4]

#uncomment these lines to use it in the IDE
#Before starting save excel files and list of candidates and ranking list in the session directory and search for URanking file somewhere. Set the following paths and patterns thereafter
candidatesFolder<-"unibz/EvaluationApp"
setwd(candidatesFolder)
excelFilesFolderPattern<-"ExcelFiles"
listCandsPattern<-"Candidates_List_2021_1"
uRankingFile<-"JANUARY2020.xlsx"

 # option_list = list(
 #   make_option(c("-d", "--folder"), type="character", default="unibz/Dropbox/SEIS/Regolamenti_BuonePratiche/EvaluationApp",
 #               help="path to cadidates' folder of excel files", metavar="character"),
 #   make_option(c("-p1", "--pattern1"), type="character", default="ExcelFiles",
 #               help="pattern to search excel files folder [default= %default]", metavar="character"),
 #   make_option(c("-p2", "--pattern2"), type="character", default="Lista_candidati_SEIS",
 #               help="pattern to search file of candidates' list [default= %default]", metavar="character"),
 #   make_option(c("-p2", "--pattern3"), type="character", default="JANUARY2020.xlsx",
 #               help="pattern to search university ranking [default= %default]", metavar="character")
 # );
 #
 # opt_parser = OptionParser(option_list=option_list);
 # opt = parse_args(opt_parser);
 
condidaciesDirectory <- file.path(getwd(), fsep = .Platform$file.sep)
print(paste("You are in the directory :", condidaciesDirectory, sep = ''))

#upload excel files, summury file, and university ranking file from folder condidaciesDirectory
candidaciesDir <- dir(condidaciesDirectory, pattern = excelFilesFolderPattern, include.dirs=T, recursive = T)
candidaciesList <- dir(paste(condidaciesDirectory, candidaciesDir, sep="/"), recursive = T)

universityRanking<-as.data.frame(read_excel(uRankingFile, sheet=1, col_names=T))

listOfCandidatesDir<-dir(condidaciesDirectory, pattern = listCandsPattern, include.dirs=T, recursive = T)
listOfCandidates<-as.data.frame(read_excel(listOfCandidatesDir[1],sheet=1, col_names=T))

#extract info on being EU or NONEU
beingNEuropean<-listOfCandidates$"EU/Non EU-citizen"
beingNEuropean[which(beingNEuropean=="ü")]<-rep("NON-EU",length(beingNEuropean[which(beingNEuropean=="ü")]))
beingNEuropean[which(is.na(beingNEuropean))]<-rep("EU",length(beingNEuropean[which(is.na(beingNEuropean))]))

#Vocabularies to verify working experience and BSc type

allowedTerms<-c("Software", "Computer", "Informatica", "Informatik", "informatiche", "Information", "dell’informazione", "developer", "development", "sviluppatore", "entwicklung")

allowedTermsBest<-c("Software Engineering", "Computer Engineering", "Computer Science", "Scienze e tecnologie informatiche",
                    "Ingegneria dell’informazione", "Information Engineering")

allowedTermsGood<-c("Scienze e tecnologie fisiche", "physics", "Scienze matematiche", "mathematics",
                    "Ingegneria industriale", "industrial engineering")

#FUNCTIONS
#Returns whether the degree type is 3, 2 or 1
checkDegree<-function(studiedTopics){
  count<-0
  degreeScore<-0
  for(i in 1:nrow(studiedTopics)){
    if(is.na(studiedTopics[i,2])==FALSE || is.na(studiedTopics[i,3])==FALSE){
      count<-count+1
    }
  }
  if(is.null(count)){
    cat(paste("No topic provided", sep=""), file = itemsToReview, append=T)
    #print("No topic provided")
  }else{
    if(count>6){
      degreeScore<-2
      cat(paste("number of courses provided is fine, manually check the score ",degreeScore, sep=""), file = itemsToReview, append=T, sep="\n")
    }else{
      if(count==6){
        degreeScore<-1
        cat(paste("number of courses provided is fine, manually check the score ",degreeScore, sep=""), file = itemsToReview, append=T, sep="\n")
      }else{
        cat(paste("number of courses provided is not sufficient, manually check the score ",degreeScore, sep=""), file = itemsToReview, append=T, sep="\n")
      }
    }
  }
  return(degreeScore)
}
#Compute total working months. Requires candidateData
computeWS<-function(wsRows,type){
  wsTotal<-0
  count<-0
  for(i in wsRows) {
    count<-count+1
    if(type[count]){
      if(!is.na(candidateData[i,2])){
        wsTotal<-wsTotal+as.numeric(candidateData[i,2])
      }
    }
  }
  return(wsTotal)
}
#compute grade in CGPA when MAX, MIN and candidate grade are numeric
computeGrade<-function(MAX,MIN,candidateGrade){
  grade<-((4-1)*(candidateGrade-MIN)/(MAX-MIN)) +1
  return(grade)
}
#convert degree grade when MAX MIN are LETTERS and compute grade in CGPA
convertLETTERSDegreeGrade<-function(MaxGrade,MinGrade,candidateGrade){
  MAX<-which(LETTERS==MaxGrade)
  MIN<-which(LETTERS==MinGrade)
  candidateNumericGrade<-which(LETTERS==candidateGrade)
  grade<-computeGrade(MAX,MIN,candidateNumericGrade)
  return(grade)
}
#convert single course grade when character and compute course grade in CGPA
convertCharacterCourseGrade<-function(originalScaleGrades,candidateGrade){
  originalScaleGrades<-na.omit(originalScaleGrades)
  MAX<-length(originalScaleGrades)
  MIN<-1
  if(candidateGrade %in% originalScaleGrades){
  candidateNumericGrade<-length(originalScaleGrades)-which(originalScaleGrades==candidateGrade)
  grade<-computeGrade(MAX,MIN,candidateNumericGrade)}else{grade<-NA}
  return(grade)
}
#convert single course grade when numeric and compute course grade in CGPA
convertCourseGrade<-function(originalScaleGrades,candidateGrade){
  originalScaleGrades<-na.omit(originalScaleGrades)
  MAX<-head(originalScaleGrades,1)
  MIN<-tail(originalScaleGrades,1)
  if(candidateGrade %in% originalScaleGrades){
  grade<-computeGrade(MAX,MIN,candidateGrade)}else{grade<-NA}
  return(grade)
}

#CANDIDATES' EVALUATION
evaluation<-data.frame()
dir.create(paste(getwd(),"Evaluation",sep="/"))

itemsToReview<-paste(getwd(),"/Evaluation/itemsToReview.txt", sep="")
if(file.exists(itemsToReview)){
  file.remove(itemsToReview)
  print("File itemsToReview.txt has been removed", sep="")
}
evaluationFile<-paste(getwd(), "/Evaluation/_",Sys.Date(), "_evaluation.csv",sep="")
if(file.exists(evaluationFile)){
  file.remove(evaluationFile)
  print("File evaluation...csv has been removed", sep="")
}


for(i in 1:length(candidaciesList)){
  candidateData <- as.data.frame(read_excel(paste(condidaciesDirectory, candidaciesDir, candidaciesList[i],sep="/"), 
                                            sheet = 1, range="B16:c29", col_names = c("DataItem", "DataValue")))
  studiedTopics <- as.data.frame(read_excel(paste(condidaciesDirectory, candidaciesDir, candidaciesList[i],sep="/"), 
                                            sheet = 2, range="B11:D24", col_names = T))
  BScProgramInfo<-as.data.frame(read_excel(paste(condidaciesDirectory, candidaciesDir, candidaciesList[i],sep="/"), 
                                           sheet = 1,  col_names = F))
  
  startF<-which(BScProgramInfo[,1]=="PASSED EXAMS")
  endF<-which(BScProgramInfo[,1]=="EXAMS TO BE TAKEN TILL GRADUATION")-1
  t<-paste("B",startF,":I",endF,sep="")
  candidateExamData <- as.data.frame(read_excel(paste(condidaciesDirectory, candidaciesDir, candidaciesList[i],sep="/"), 
  
  sheet = 1, range=t, col_names = F))
  
  GD<-"(Expected) Date of graduation"
  graduationDate<-as.Date(as.character(BScProgramInfo[which(BScProgramInfo[,2]==GD)+1,2]), "%d.%m.%y")
  
  bscD<-"STANDARD DURATION OF PREVIOUS BACHELOR COURSE:"
  bscDuration<-as.numeric(str_split(candidateData[which(candidateData[,1]==bscD),2], pattern=" ")[[1]][1])
  
  bscName<-"NAME OF PREVIOUS BACHELOR COURSE:"
  bscRows<-which(candidateData[,1]==bscName)
  bsc<-candidateData[bscRows,2]
  
  university<-"UNIVERSITY AWARDING THE STUDY TITLE:"
  universityRow<-which(candidateData[,1]==university)
  universityName<-candidateData[universityRow,2]
  
  url<-"WEBSITE OF THE UNIVERSITY:"
  universityRow<-which(candidateData[,1]==url)
  universityURL<-candidateData[universityRow,2]
  
  ws<-"WORKING EXPERIENCE (in months)"
  wsRows<-which(candidateData[,1]==ws)
  wsType<-wsRows+1
  
  #weights default values
  wStud<-0.6
  wBSc<-0.3
  wWE<-0.1
  wCV<-1
  CV<-5
  scoresWeights<-c(wStud,wBSc,wWE,wCV)
 
  cat(paste("CANDIDATE: ",candidateData[1,2]," ", candidateData[2,2], sep=""), file = itemsToReview, append=T, sep="\n")
  #print(paste("CANDIDATE: ",candidateData[1,2]," ", candidateData[2,2], sep=""))
  
  indexSurname<-which(unlist((lapply(str_split(listOfCandidates$Lastname, pattern=" "),function(x){candidateData[1,2] %in% x}))))
  indexSurnameSwap<-which(unlist((lapply(str_split(listOfCandidates$Lastname, pattern=" "),function(x){candidateData[2,2] %in% x}))))
  indexName<-which(unlist((lapply(str_split(listOfCandidates$Firstname, pattern=" "),function(x){candidateData[2,2] %in% x}))))
  indexNameSwap<-which(unlist((lapply(str_split(listOfCandidates$Firstname, pattern=" "),function(x){candidateData[1,2] %in% x}))))
 if(length(indexSurname)==0 & length(indexName)==0 &length(indexSurnameSwap)==0 & length(indexNameSwap)==0){
   cat(paste("Candidate not found", sep=""), file = itemsToReview, append=T, sep="\n")
   beingNEU<-"Not found"
 }else{
   if(length(indexSurnameSwap)>0 || length(indexNameSwap)>0){
     cat(paste("name surname have been swapped: check name",sep=""), file = itemsToReview, append=T, sep="\n")
     if(length(indexSurnameSwap)>0){
       beingNEU<-beingNEuropean[indexSurnameSwap]}else{
       beingNEU<-beingNEuropean[indexNameSwap]
     }
   }else{
   if(length(indexSurname)==0){
  beingNEU<-beingNEuropean[indexName]
   }else{
     beingNEU<-beingNEuropean[indexSurname]
   }
 }}
  
  #OUTPUT
  #Type of BSc degree. Default 0
  degreeScore<-0
  #number of months of work. Default 0
  wsCount<-0
  #Still to graduate? Default False
  dateScore<-F
  #Duration of the BSc. Default empty char
  bscDurationFlag<-""
  #degree grade in CGPA. DEfualt is 0
  weightedGrade<-0
  #final score of he candidate. 6 to pass
  score<-5
  
  #Evaluates the type of BSc
  typeBestBoolean<-str_detect(bsc, coll(allowedTermsBest,ignore_case=T), negate = F)
  typeGoodBoolean<-str_detect(bsc, coll(allowedTermsGood,ignore_case=T), negate = F)
  if(any(typeBestBoolean)){
    degreeScore<-3
    #print(paste("Degree is eligible: ", bsc,"; scoreBScType ",degreeScore, sep=""))
  }else{
    if(any(typeGoodBoolean) & checkDegree(studiedTopics)>0){
      degreeScore<-checkDegree(studiedTopics)
      #print(paste("Degree is eligible: ", bsc,"; score BSc type ",degreeScore, sep=""))
    }else{
      cat(paste("Degree is not eligible", sep=""), file = itemsToReview,  append=T, sep="\n")
    }
  }
  
  #creates the terms from type of working experience and verifies whether they concern the program field
  termType<-list()
  type<-c(F,F,F)
  count<-0
  for(j in wsType){
    count<-count+1
    termType[[count]]<-tolower(unlist(str_split(candidateData[j,2],pattern=" ")))
    if(any(termType[[count]] %in% tolower(allowedTerms))){
      type[count]<-T
      if(!all(type)){
      cat(paste("Check working experience", sep=" "), file = itemsToReview, append=T, sep="\n")}
      #print(cat("Working experience suitable as ",termType[[count]], sep=" "))
    }
  }
  
  #evaluates the score for working experience only for suitable ws
  
  wsMonths<-computeWS(wsRows,type)
  if(wsMonths>=36){wsCount<-3}else{if(wsMonths>=24){wsCount<-2}else{if(wsMonths>=12){wsCount<-1}}}
  cat(paste("WORKING EXPERIENCE: ", wsMonths, " score: ", wsCount, sep=""), file = itemsToReview, append=T, sep="\n")
  #print(paste("WORKING EXPERIENCE: ", wsMonths, " score: ", wsCount, sep=""))
  
  #verifies whether the date of graduation occurs in the future.
  if(is.na(graduationDate) || as.Date(graduationDate)>Sys.Date()){
    dateScore<-T
  }
  
  #evaluates whether the BSc duration is legally fine.
  if(is.na(bscDuration)){
    bscDurationFlag<-"not given"
    cat(paste("Duration not given",sep=""), file = itemsToReview, append=T, sep="\n")
    #print("Duration not OK")
  }else{
    if(bscDuration>2){
      bscDurationFlag<-"OK"
      #print("Duration OK")
    }else{ 
      bscDurationFlag<-"not OK"
      cat(paste("Duration not OK", sep=""), file = itemsToReview, append=T, sep="\n")
      #print("Duration not OK")
    }
  }
  #compute grade conversions: grade in in CGPA
  candidateGrade<-candidateExamData[which(candidateExamData[,6]=="Your Grade:")+1,6]
  MaxGrade<-candidateExamData[which(candidateExamData[,6]=="Maximum Grade:")+1,6]
  MinGrade<-candidateExamData[which(candidateExamData[,6]=="Pass Grade:")+1,6]
  grade<-0
  numcandidateGrade<-as.numeric(candidateGrade)
  numMax<-as.numeric(MaxGrade)
  numMin<-as.numeric(MinGrade)
  if(all(!is.na(numcandidateGrade),!is.na(numMax),!is.na(numMax))){
    grade<-computeGrade(numMax,numMin,numcandidateGrade)
  }else{
    if(is.na(candidateGrade)|| !(any(LETTERS==MaxGrade, na.rm=T)&any(LETTERS==MinGrade, na.rm=T)&any(LETTERS==candidateGrade, na.rm=T))){
      start<-which(candidateExamData=="PASSED EXAMS")
      end<-nrow(candidateExamData)
      coursesList<-candidateExamData[(start+2):end,-c(5,7)]
      names(coursesList)<-c(candidateExamData[3,-c(5:8)],candidateExamData[2,c(6,8)])
      temp<-coursesList$"GRADING SCALE 1"
      originalScaleGrades<-na.omit(toupper(temp[(which(temp=="Full Grading Scale:")+1):length(temp)]))
      candidateGrades<-na.omit(toupper(coursesList$Grade))
      if(is.character(originalScaleGrades)){
        courseGrades<-unlist(lapply(candidateGrades,function(x){convertCharacterCourseGrade(originalScaleGrades,x)}))
      }else{
        courseGrades<-unlist(lapply(candidateGrades,function(x){convertCourseGrade(originalScaleGrades,x)}))
      }
      coursesCredits<-as.numeric(na.omit(coursesList$Credits))
      grade<-sum(coursesCredits*courseGrades)/sum(coursesCredits)
    }else{
      grade<-convertLETTERSDegreeGrade(MaxGrade,MinGrade,candidateGrade)
    }
  }

  #computes the weighted grade. Use the last veriosn of the University ranking
  universityRankingScore<-0
  #universityName<-"Chapman University"
  #universityURL<- "http://www.chapman.edu/"
  #universityRanking$WR[which.max(lapply(universityRanking$NAME, function(x) sum(universityName %in% x)))]
  #rankingNAME<-universityRanking$WR[which(any(universityRanking$NAME %in% str_split(universityName, pattern=" ")))]
  rankingNAME<-universityRanking$WR[[which.max(lapply(universityRanking$NAME, function(x){
    sum(unlist(str_split(universityName, pattern=" ")) %in% unlist(str_split(x, pattern=" ")))
  }))]]
  rankingURL<-universityRanking$WR[which(universityRanking$URL==universityURL)]
  if(length(rankingNAME)==0 & length(rankingURL)==0){
    cat(paste("not able to find university ranking of ",universityName,sep=""), file = itemsToReview, append=T, sep="\n")
    #print(paste("not able to find university ranking of ",universityName,sep=""))
  }else{
    if(length(rankingNAME)==0){ranking<-rankingURL}else{ranking<-rankingNAME}
  if(ranking>4000){
    universityRankingScore=0
  }else{
    if(ranking>3000){
      universityRankingScore=0.5
    }else{
      if(ranking>2000){
        universityRankingScore=1
      }else{
        if(ranking>1000){
          universityRankingScore=2
        }else{
          universityRankingScore=3}}}}
  }
  weightedGrade<-grade*universityRankingScore

  score<-(weightedGrade*wStud+degreeScore*wBSc+wsCount*wWE)+CV*wCV
  result<-NULL
  if(score>=6){
    #print("candidate is admitted")
    result<-"admitted"
  }else{
    cat(paste("candidate is not admitted",sep=""),file = itemsToReview, append=T, sep="\n")
    #print("candidate is not admitted")
    result<-"not admitted"
  }
  
  candidatedf<-data.frame(NAME = candidateData[1,2], 
                          SURNAME = candidateData[2,2], 
                          BEING_NEUROPEAN = beingNEU,
                          DURATION_BSC = bscDuration,
                          WEIGHTED_GRADE = weightedGrade,
                          TYPE_BSC= degreeScore, 
                          WORKING_EXPERIENCE_MONTHS = wsCount, 
                          TOTAL_SCORE=score, 
                          RESULT= result,
                          STILL_TO_GRADUATE= dateScore
                          )
  
  evaluation<-rbind(evaluation, candidatedf)
}

write.csv(evaluation, file=evaluationFile, append=T)

line<-paste("Evaluation ",Sys.time(),sep="")
cat(line,file=evaluationFile, append=T)
cat(line,file=itemsToReview, append=T)
options(warn = defaultW)
