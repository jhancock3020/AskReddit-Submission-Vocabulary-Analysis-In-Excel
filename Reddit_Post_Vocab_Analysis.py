import praw
import xlwt
import xlrd
from xlutils.copy import copy
from xlwt import Workbook
from datetime import datetime

#REDDIT SETUP SECTION
#Access Reddit API
reddit = praw.Reddit(client_id='14 Char Info', 
                     client_secret='24 Char Info', 
                     user_agent='Project Name')
#The subreddit object that will be analyzed
askReddit = reddit.subreddit('askreddit')
subLimit = 35
#Gets the top posts from the subreddit
askRedditTopPosts = askReddit.hot(limit=subLimit)
#Gets the position (in the context of sheet placement) of newly collected posts
redditRank = 1

#EXCEL WORKBOOK SECTION
#Opens and reads previously manually created excel file in the form of two sheets
rb = xlrd.open_workbook('fileExcel.xls',formatting_info=True)
readSheet = rb.sheet_by_index(0)
readSheet1 = rb.sheet_by_index(1)
#copies what is read to a new book you can write to, along with the same two sheets inside said book
wb = copy(rb)
writeSheet = wb.get_sheet(0)
writeSheet1 = wb.get_sheet(1)
writeSheet2 = wb.get_sheet(2)
writeSheet3 = wb.get_sheet(3)
writeSheet4 = wb.get_sheet(4)
#FONT SECTION
#different font styles used
styleBold = xlwt.easyxf('font: bold 1, color black;')
styleHigh = xlwt.easyxf('font: bold 1, color red;')
styleLow = xlwt.easyxf('font: bold 1, color blue;')
styleBasic = xlwt.easyxf('font: bold off, color black;')

#MISC SECTION
#removeChar chars that should not be read as part of the input of words
removeChar = ".!,'’‘?""-"
#Gets the current date
today = datetime.now()
currDate = today.strftime("%d/%m/%Y %I:%M:%S %p")
#list of indivual words
wordList = []
#number of times those words appear
wordCount = []
#List of prevously recorded submissions
submissionList = []

#if workbook is empty, add these as first row titles of columns
if(readSheet.nrows <= 0):
    writeSheet.write(0,0, "WORD LIST",styleBold)
    writeSheet.write(0,1, "Total Num",styleBold)
    writeSheet1.write(0,0,"Post ID",styleBold)
    writeSheet1.write(0,1,"Post Title",styleBold)
    writeSheet1.write(0,2,"Author",styleBold)
    writeSheet1.write(0,3,"# Of Comments",styleBold)
    writeSheet1.write(0,4,"# Of Upvotes",styleBold)
    writeSheet1.write(0,5,"Upvote Ratio",styleBold)
    writeSheet1.write(0,6,"Over 18",styleBold)
    writeSheet1.write(0,7,"Serious Tag",styleBold)
    writeSheet1.write(0,8,"# Of Plat Awards",styleBold)
    writeSheet1.write(0,9,"# Of Gold Awards",styleBold)
    writeSheet1.write(0,10,"# Of Silver Awards",styleBold)
    writeSheet1.write(0,11,"Was Updated",styleBold)   
#Save and reopen the workbook and update the readSheets
wb.save('fileExcel.xls') 
rb = xlrd.open_workbook('fileExcel.xls')
readSheet = rb.sheet_by_index(0)
readSheet1 = rb.sheet_by_index(1)

#startCol is the next blank column where currDate and new instances of words can be added
startCol = readSheet.ncols
#If the word list in excel has words in it (word list doesnt include the WORD LIST title at 0,0), add it to the python list and append the curr day's count to 0 until proven otherwise
if readSheet.nrows > 0:
    for i in range(readSheet.nrows):
        if(i != 0):
            wordList.append(readSheet.cell_value(i,0))
            wordCount.append(0)
#If the submission list in excel has submissions in it, add the ID to the python list          
if readSheet1.nrows > 0:
    for i in range(readSheet1.nrows):
        submissionList.append(readSheet1.cell_value(i,0))
#Go through the current top reddit posts
for j,submission in enumerate(askRedditTopPosts):
    if (submission.stickied == True):
        subLimit += 1
askRedditTopPosts = askReddit.hot(limit=subLimit)

#MAIN LOOP
#Go through the current top reddit posts
for j,submission in enumerate(askRedditTopPosts):
    #If the curr sub's ID is not in the list of previously recorded submissions
    if(submission.id not in submissionList and submission.stickied == False):
        print(str(redditRank)+ " " + submission.title +"\n")
        #Add sub ID to sub ID list
        submissionList.append(submission.id)

        #READ SUBMISSION TITLE AND COUNT INDIVIDUAL WORDS IN WORDLIST
        #Create a string variable of sub's title and remove unwanted chars 
        editedTitle = submission.title
        editedTitle = editedTitle.lower()
        for character in removeChar:
            editedTitle = editedTitle.replace(character,"")
        editedTitle = editedTitle.replace("{serious}","")
        editedTitle = editedTitle.replace("[serious]","")
        editedTitle = editedTitle.replace("(serious)","")
        editedTitle = editedTitle.replace("{nsfw}","")
        editedTitle = editedTitle.replace("[nsfw]","")
        editedTitle = editedTitle.replace("(nsfw)","")  
        #Then make all letters in the string lowercase before spliting it into a list of strings split by spaces
        currTitle = editedTitle.split()
        #go through char list of title words, if the word has already been used, add 1 to the corresponding wordCount index
        for word in currTitle:
            for i,wordInList in enumerate(wordList):
                if word == wordInList:
                    wordCount[i] += 1
            #If it hasnt been seen before, add it to the word list and give it a count of 0.
            if word not in wordList:
                wordList.append(word)
                wordCount.append(1)

        #WRITE SUBMISSION INFO TO SHEET
        #Write current date on next blank row
        writeSheet1.write(readSheet1.nrows,0,currDate,styleBold)
        #Writes relevant data to second sheet about each submission and increment the rank
        writeSheet1.write(readSheet1.nrows+redditRank,0,submission.id)
        writeSheet1.write(readSheet1.nrows+redditRank,1,submission.title)
        #if sub author name is deleted, mark it as such. Otherwise, just put the name
        if submission.author is None:
            author = "[DELETED]"
        else:
            author = submission.author.name
        writeSheet1.write(readSheet1.nrows+redditRank,2,author)
        writeSheet1.write(readSheet1.nrows+redditRank,3,submission.num_comments)
        writeSheet1.write(readSheet1.nrows+redditRank,4,submission.score)
        writeSheet1.write(readSheet1.nrows+redditRank,5,submission.upvote_ratio)
        writeSheet1.write(readSheet1.nrows+redditRank,6,submission.over_18)
        #if submission has a Serious tag, mark it as being True, otherwise False
        if submission.link_flair_text == 'Serious Replies Only':
           writeSheet1.write(readSheet1.nrows+redditRank,7,bool(1),styleHigh)
        else:
           writeSheet1.write(readSheet1.nrows+redditRank,7,bool(0),styleLow)
        #Goes through different possible awards and checks to see if current submission has any, marks them if they do
        numSilver = submission.gildings.get("gid_1")
        if numSilver is None:
            numSilver = 0
        numGold = submission.gildings.get("gid_2")
        if numGold is None:
            numGold = 0
        numPlat = submission.gildings.get("gid_3")
        if numPlat is None:
            numPlat = 0
        writeSheet1.write(readSheet1.nrows+redditRank,8,numPlat)
        writeSheet1.write(readSheet1.nrows+redditRank,9,numGold)
        writeSheet1.write(readSheet1.nrows+redditRank,10,numSilver)
        
        #iterates redditRank by +1
        redditRank += 1
    #If ID is in list
    elif(submission.id in submissionList):
        #Go through the rows
        updateInfo = ""
        for i in range(readSheet1.nrows):
            #If ID matches on in list, check to see if any of the numeric values changed and update them on sheet
            if(submission.id == readSheet1.cell_value(i,0)):
                if(submission.num_comments != readSheet1.cell_value(i,3)):
                    updateInfo += "COMMENTS from " + str(int(readSheet1.cell_value(i,3)))+" to " + str(submission.num_comments) + " "
                    writeSheet1.write(i,3,int(submission.num_comments))
                if(submission.score != readSheet1.cell_value(i,4)):
                    updateInfo += "UPVOTES from " + str(int(readSheet1.cell_value(i,4)))+" to " + str(submission.score) + " "
                    writeSheet1.write(i,4,submission.score)
                if(submission.upvote_ratio != readSheet1.cell_value(i,5)):
                    updateInfo += "RATIO from " + str(readSheet1.cell_value(i,5))+" to " + str(submission.upvote_ratio) + " "
                    writeSheet1.write(i,5,submission.upvote_ratio)
                #Goes through different possible awards and checks to see if current submission has any, marks them if they do
                #If there are no awards, set award number to zero
                numPlat = submission.gildings.get("gid_3")
                numGold = submission.gildings.get("gid_2")
                numSilver = submission.gildings.get("gid_2")
                if numSilver is None:
                    numSilver = 0
                numGold = submission.gildings.get("gid_2")
                if numGold is None:
                    numGold = 0
                numPlat = submission.gildings.get("gid_3")
                if numPlat is None:
                    numPlat = 0
                if(numPlat > readSheet1.cell_value(i,8)):
                    updateInfo += "PLAT from " + str(readSheet1.cell_value(i,8))+" to " + str(submission.gildings.get("gid_3")) + " "
                    writeSheet1.write(i,8,submission.gildings.get("gid_3"))
                if(numGold > readSheet1.cell_value(i,9)):
                    updateInfo += "GOLD from " + str(readSheet1.cell_value(i,9))+" to " + str(submission.gildings.get("gid_2")) + " "
                    writeSheet1.write(i,9,submission.gildings.get("gid_2"))
                if(numSilver > readSheet1.cell_value(i,10)):
                    updateInfo += "SILV from " + str(readSheet1.cell_value(i,10))+" to " + str(submission.gildings.get("gid_1")) + " "
                    writeSheet1.write(i,10,submission.gildings.get("gid_1"))
                #If content was updated, write info about it to the las tcell in the row
                if(updateInfo != ""):
                    updateInfo += "was updated on " + currDate
                    writeSheet1.write(i,11,updateInfo)

#WRITE WORD LIST AND CURR WORD COUNT TO SHEET
if all([ v != 0 for v in wordCount ]) :
    print('indeed they are')
#Writes the current date and time of when submissions were received on wordSheet's first row above the current word count list        
writeSheet.write(0, startCol, currDate,styleBold)
#Rewrites the word list into cells and writes current day's word count to the next blank row
for i,word in enumerate(wordList):
    writeSheet.write(i+1, 0, word)
    writeSheet.write(i+1, startCol, wordCount[i])
#Save and reopen the workbook and update the readSheets
wb.save('fileExcel.xls') 
rb = xlrd.open_workbook('fileExcel.xls',formatting_info=True)
readSheet = rb.sheet_by_index(0)
readSheet1 = rb.sheet_by_index(1)
readSheet2 = rb.sheet_by_index(2)

#If an empty cell exists in the word count sheet, make it a zero
for i in range(readSheet.nrows):
    for j in range(readSheet.ncols):
        if(readSheet.cell_value(i,j) == ''):
            writeSheet.write(i,j,0)
#Save and reopen the workbook and update the readSheets
wb.save('fileExcel.xls') 
rb = xlrd.open_workbook('fileExcel.xls',formatting_info=True)
readSheet = rb.sheet_by_index(0)
readSheet1 = rb.sheet_by_index(1)
readSheet2 = rb.sheet_by_index(2)
readSheet3 = rb.sheet_by_index(3)

#Goes through each row and add the total number of each word's wordCount instance
for i in range(readSheet.nrows):
    #total count of all instances of a word
    tot = 0
    for j in range(readSheet.ncols):
        if(i > 0 and j > 1):
            tot += readSheet.cell_value(i,j)
    if(i != 0):
        writeSheet.write(i,1, tot,styleBold)
#Save and reopen the workbook and update the readSheets
wb.save('fileExcel.xls') 
rb = xlrd.open_workbook('fileExcel.xls',formatting_info=True)
readSheet = rb.sheet_by_index(0)
readSheet1 = rb.sheet_by_index(1)
readSheet2 = rb.sheet_by_index(2)
readSheet3 = rb.sheet_by_index(3)

#EDITS TEXT COLOR AND FORMAT OF HIGHEST AND LOWEST NUMBERS OF SUBMISSION INFO
#Go through the different columns to find the highest and lowest numbers in each
for i in range(readSheet1.ncols):
    low = 1000000
    high = 0
    for j in range(readSheet1.nrows):
        if(j != 0 and i > 2 and i != 6 and i != 7 and i != 11 and readSheet1.cell_value(j,i) != ''):
            if(readSheet1.cell_value(j,i) < low):
                low = readSheet1.cell_value(j,i)
            if(readSheet1.cell_value(j,i) > high):
                high = readSheet1.cell_value(j,i)
    #Per column, if number is highest (or bool is True), bold and make red.
    #If number is lowest (or bool is False), bold and make blue
    #Else, just make it regular font.   
    for j in range(readSheet1.nrows):
        if(j != 0 and i > 2 and i != 6 and i != 7 and i != 11 and readSheet1.cell_value(j,i) != ''):
            if(readSheet1.cell_value(j,i) == low):
                writeSheet1.write(j,i,low,styleLow)
            elif(readSheet1.cell_value(j,i) == high):
                writeSheet1.write(j,i,high,styleHigh)
            else:
                writeSheet1.write(j,i,readSheet1.cell_value(j,i),styleBasic)
        elif (j != 0 and  (i == 6 or i == 7) and readSheet1.cell_value(j,i) != ''):
            if(readSheet1.cell_value(j,i) == 0):
                writeSheet1.write(j,i,bool(readSheet1.cell_value(j,i)),styleLow)
            elif(readSheet1.cell_value(j,i) == 1):
                writeSheet1.write(j,i,bool(readSheet1.cell_value(j,i)),styleHigh)
data = []
cols2 = 0
#Fill in 2d list. Each element contains all info on submission
for j in range(readSheet1.nrows):
    elem = []
    if(j != 0 and readSheet1.cell_value(j,2) != ''):
        elem = [readSheet1.cell_value(j,i)for i in range(readSheet1.ncols)]
        data.append(elem)
#Goes though each column from sheet1 and sorts questions from highest to lowest
for i in range(11):
    if(i > 2 and i != 6 and i != 7 and i != 11):
        data = sorted(data,key=lambda l:l[i],reverse=True)
        writeSheet2.write(0,cols2,readSheet1.cell_value(0,1),styleBold)
        writeSheet2.write(0,cols2+1,readSheet1.cell_value(0,i),styleBold)
        for j in range(len(data)):
            writeSheet2.write(1+j,cols2,data[j][1])
            writeSheet2.write(1+j,cols2+1,data[j][i])
        cols2 += 3
data = []
cols1 = 0
#Fill in 2d list. Each element contains all info on submission
for j in range(readSheet.nrows):
    if(j != 0):
        elem = [readSheet.cell_value(j,i)for i in range(readSheet.ncols)]
        data.append(elem)
for j in range(readSheet.ncols):
    if(j > 0):
        data = sorted(data,key=lambda x:x[j],reverse=True)
        num = int(cols1/3)+1
        for i in range(len(data)):
            writeSheet3.write(0,cols1,readSheet.cell_value(0,num),styleBold)
            writeSheet3.write(1+i,cols1,data[i][0])
            writeSheet3.write(1+i,cols1+1,data[i][j])
        cols1 += 3
        
#Save and reopen the workbook and update the readSheets
wb.save('fileExcel.xls') 
rb = xlrd.open_workbook('fileExcel.xls',formatting_info=True)
readSheet = rb.sheet_by_index(0)
readSheet1 = rb.sheet_by_index(1)
readSheet2 = rb.sheet_by_index(2)
readSheet3 = rb.sheet_by_index(3)
readSheet4 = rb.sheet_by_index(4)
#Arrays to hold newly made word list, question list, and comments/upvotes of those questions
wordData = []
questionData = []
comments = []
points = []
#Keeps track of where to write word and its questions in overall sheet
keepTackPos = 1
#Read word list from total most used word to least used word and add it to a word list
for j in range(readSheet3.nrows):
    if(j > 0):
        wordData.append(readSheet3.cell_value(j,0))
#Go through question data
for j in range(readSheet1.nrows):
    if(j > 0 and readSheet1.cell_value(j,1)!= ''):
        #clean up question words by getting rid of unwanted chars
        #and make it all lowercase to make mathcing easy
        sentence = readSheet1.cell_value(j,1).lower()
        for character in removeChar:
            sentence = sentence.replace(character,"")
        #add question, its upvotes, and its comments to their own lists.
        questionData.append(sentence)
        comments.append(readSheet1.cell_value(j,3))
        points.append(readSheet1.cell_value(j,4))
writeSheet4.write(0,0,"Words and Questions",styleBold)
writeSheet4.write(0,1,"Upvotes",styleBold)
writeSheet4.write(0,2,"Comments",styleBold)
#Go through all words in wordlist
for j in range(len(wordData)):
    #Write word from list
    writeSheet4.write(keepTackPos,0,wordData[j],styleBold)
    #Create variables for each word's respective total comments and upvotes
    currPoint = 0
    currComments = 0
    #Keep track of row where word is located, so total comments and upvotes can be recorded
    wordTitlePos = keepTackPos
    #Go through all the questions
    for i in range(len(questionData)):
        #If the current word is in the question
        if(wordData[j] in questionData[i]):
            #Write question, its upvotesd, and its comments
            keepTackPos += 1
            writeSheet4.write(keepTackPos,0,questionData[i])
            writeSheet4.write(keepTackPos,1,points[i])
            writeSheet4.write(keepTackPos,2,comments[i])
            currPoint += points[i]
            currComments += comments[i]
    #Write the total upvotes and comments next to word
    writeSheet4.write(wordTitlePos,1,currPoint,styleBold)
    writeSheet4.write(wordTitlePos,2,currComments,styleBold)
    keepTackPos += 1
        
wb.save('fileExcel.xls') 
