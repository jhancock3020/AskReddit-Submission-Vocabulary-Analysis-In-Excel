import praw
import xlwt
import xlrd
from xlutils.copy import copy
from xlwt import Workbook
from datetime import datetime

#Access Reddit API
reddit = praw.Reddit(client_id='14 Char Info', 
                     client_secret='24 Char Info', 
                     user_agent='Project Name')
#The subreddit object that will be analyzed
askReddit = reddit.subreddit('askreddit')
subLimit = 10
#Gets the top posts from the subreddit
askRedditTopPosts = askReddit.hot(limit=subLimit)
#Gets the position (in the context of sheet placement) of newly collected posts
redditRank = 1
#Opens and reads previously manually created excel file in the form of two sheets
rb = xlrd.open_workbook('fileExcel.xls',formatting_info=True)
readSheet = rb.sheet_by_index(0)
readSheet1 = rb.sheet_by_index(1)
#copies what is read to a new book you can write to, along with the same two sheets inside said book
wb = copy(rb)
writeSheet = wb.get_sheet(0)
writeSheet1 = wb.get_sheet(1)
#removeChar chars that should not be read as part of the input of words
removeChar = ".!,'’‘?"""
#Gets the current date
today = datetime.now()
currDate = today.strftime("%d/%m/%Y %I:%M:%S %p")
#list of indivual words
wordList = []
#number of times those words appear
wordCount = []
#List of prevously recorded submissions
submissionList = []
#different font styles used
styleBold = xlwt.easyxf('font: bold 1, color black;')
styleHigh = xlwt.easyxf('font: bold 1, color red;')
styleLow = xlwt.easyxf('font: bold 1, color blue;')
styleBasic = xlwt.easyxf('font: bold off, color black;')
#if workbook is empty, add these as first row titles of columns
if(readSheet.nrows <= 0):
    writeSheet.write(0,0, "WORD LIST",styleBold)
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
#Go through the current top reddit posts
for j,submission in enumerate(askRedditTopPosts):
    #If the curr sub's ID is not in the list of previously recorded submissions
    if(submission.id not in submissionList and submission.stickied == False):
        print(str(redditRank)+ " " + submission.title +"\n")
        #Create a string variable of sub's title and remove unwanted chars 
        editedTitle = submission.title
        for character in removeChar:
            editedTitle = editedTitle.replace(character,"")
        editedTitle = editedTitle.replace("[Serious]","")        
        #Then make all letters in the string lowercase before spliting it into a list of strings split by spaces
        currTitle = editedTitle.lower().split()
        #go through char list of title words, if the word has already been used, add 1 to the corresponding wordCount index
        for word in currTitle:
            for i,wordInList in enumerate(wordList):
                if word == wordInList:
                    wordCount[i] += 1
            #If it hasnt been seen before, add it to the word list and give it a count of 0.
            if word not in wordList:
                wordList.append(word)
                wordCount.append(1)
        #Add sub ID to sub ID list
        submissionList.append(submission.id)
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
#Writes the current date and time of when submissions were received on wordSheet's first row above the current word count list        
writeSheet.write(0, startCol, currDate+" # of instances of word on this date and time",styleBold)
#Rewrites the word list into cells and writes current day's word count to the next blank row
for i,word in enumerate(wordList):
    writeSheet.write(i+1, 0, word)
    writeSheet.write(i+1, startCol, wordCount[i])
#Save and reopen the workbook and update the readSheets
wb.save('fileExcel.xls') 
rb = xlrd.open_workbook('fileExcel.xls',formatting_info=True)
readSheet = rb.sheet_by_index(0)
readSheet1 = rb.sheet_by_index(1)
#If an empty cell exists in the word count sheet, make it a zero
for i in range(readSheet.nrows):
    for j in range(readSheet.ncols):
        if(readSheet.cell_value(i,j) == ''):
            writeSheet.write(i,j,0)
low = 100
high = 0
low2 = 1000000
high2 = 0
low3 = 1000000
high3 = 0
lowGold = 1000000
highGold = 1
lowSil = 1000000
highSil = 1
lowPlat = 1000000
highPlat = 1
#Go through the different columns to find the highest and lowest numbers in each
for i in range(readSheet1.nrows):
    if(i != 0):
        if(readSheet1.cell_value(i,5) != ''):
           if(readSheet1.cell_value(i,5) < low):
               low = readSheet1.cell_value(i,5)
           if(readSheet1.cell_value(i,5) > high):
               high = readSheet1.cell_value(i,5)
        if(readSheet1.cell_value(i,4) != ''):
           if(readSheet1.cell_value(i,4) < low2):
               low2 = readSheet1.cell_value(i,4)
           if(readSheet1.cell_value(i,4) > high2):
               high2 = readSheet1.cell_value(i,4)   
        if(readSheet1.cell_value(i,3) != ''):
           if(readSheet1.cell_value(i,3) < low3):
               low3 = readSheet1.cell_value(i,3) 
           if(readSheet1.cell_value(i,3) > high3):
               high3 = readSheet1.cell_value(i,3)
        if(readSheet1.cell_value(i,9) != ''):
           if(readSheet1.cell_value(i,9) < lowGold):
               lowGold = readSheet1.cell_value(i,9) 
           if(readSheet1.cell_value(i,9) > highGold):
               highGold = readSheet1.cell_value(i,9)   
        if(readSheet1.cell_value(i,10) != ''):
           if(readSheet1.cell_value(i,10) < lowSil):
               lowSil = readSheet1.cell_value(i,10) 
           if(readSheet1.cell_value(i,10) > highSil):
               highSil = readSheet1.cell_value(i,10)
        if(readSheet1.cell_value(i,8) != ''):
           if(readSheet1.cell_value(i,8) < lowPlat):
               lowSil = readSheet1.cell_value(i,8) 
           if(readSheet1.cell_value(i,8) > highPlat):
               highSil = readSheet1.cell_value(i,8)
#Per column, if number is highest (or bool is True), bold and make red.
#If number is lowest (or bool is False), bold and make blue
#Else, just make it regular font.        
for i in range(readSheet1.nrows):
    if(i != 0):
        if(readSheet1.cell_value(i,6) == 0):
            writeSheet1.write(i,6,bool(readSheet1.cell_value(i,6)),styleLow)
        elif(readSheet1.cell_value(i,6) == 1):
            writeSheet1.write(i,6,bool(readSheet1.cell_value(i,6)),styleHigh)
            
        if(readSheet1.cell_value(i,5) == low):
            writeSheet1.write(i,5,low,styleLow)
        elif(readSheet1.cell_value(i,5) == high):
            writeSheet1.write(i,5,high,styleHigh)
        else:
            writeSheet1.write(i,5,readSheet1.cell_value(i,5),styleBasic)
            
        if(readSheet1.cell_value(i,4) == high2):
            writeSheet1.write(i,4,high2,styleHigh)
        elif(readSheet1.cell_value(i,4) == low2):
            writeSheet1.write(i,4,low2,styleLow)
        else:
            writeSheet1.write(i,4,readSheet1.cell_value(i,4),styleBasic)
            
        if(readSheet1.cell_value(i,3) == high3):
            writeSheet1.write(i,3,high3,styleHigh)
        elif(readSheet1.cell_value(i,3) == low3):
            writeSheet1.write(i,3,low3,styleLow)
        else:
            writeSheet1.write(i,3,readSheet1.cell_value(i,3),styleBasic)

        if(readSheet1.cell_value(i,8) == highPlat):
            writeSheet1.write(i,8,highPlat,styleHigh)
        elif(readSheet1.cell_value(i,8) == lowPlat):
            writeSheet1.write(i,8,lowPlat,styleLow)
        else:
            writeSheet1.write(i,8,readSheet1.cell_value(i,8),styleBasic)
            
        if(readSheet1.cell_value(i,9) == highGold):
            writeSheet1.write(i,9,highGold,styleHigh)
        elif(readSheet1.cell_value(i,9) == lowGold):
            writeSheet1.write(i,9,lowGold,styleLow)
        else:
            writeSheet1.write(i,9,readSheet1.cell_value(i,9),styleBasic)
            
        if(readSheet1.cell_value(i,10) == highSil):
            writeSheet1.write(i,10,highSil,styleHigh)
        elif(readSheet1.cell_value(i,10) == lowSil):
            writeSheet1.write(i,10,lowSil,styleLow)
        else:
            writeSheet1.write(i,10,readSheet1.cell_value(i,10),styleBasic)

wb.save('fileExcel.xls') 
