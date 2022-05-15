# AskReddit-Submission-Vocabulary-Analysis-In-Excel
It is a commonly held belief that the same questions are posted onto the subreddit AskReddit on a regular basis in order to gain the most 'karma' (upvotes), with the only major difference being how each question is phrased.

The goal of this project is to use Python 3 and the PRAW library to find the top 25 (front page) posts on AskReddit for a certain period of time.
During this time, the words used in each post will be counted in order to determine the most commonly used words and phrases.
The number of comments, upvotes, upvote ratio, if the submission is marked as NSFW or 'Serious', and awards earned (among possible other attributes) will also be tracked for each submission.

**CURRENT FEATURES:**

* Tracks the number of times a word appears and displays the total number and number per program execution in columns.

* Stores each question and its attributes, including the number of comments, upvotes, upvote ratio, if the submission is marked as NSFW or 'Serious', and awards (Platinum, Gold, and Silver) earned.

* Updates information about currently stored question (ex: see if upvotes or awards have increased since the last time the question was looked at), and marks what about the question was updated.

* Ignores pinned posts

* Analyzes and dedicated sheet displays the different attributes (and their questions) in order from greatest to least.

* Analyzes and dedicated sheet displays the different word count column in order from greatest to least.

* Analyzes the word data. Ex: If the word is "dog", then lists all questions containing that word and the number of upvotes and comments each question recieved. Also displays the total number of upvotes and comments all questions containing that word combined had.


**TO DO:** 

* Format columns so they are the length of longest string

* Add data analysis of author data (ex: who posts the most, what questions do they post, how many total upvotes do they have).
