
import requests
import newspaper
from newspaper import Article
from nltk.sentiment.vader import SentimentIntensityAnalyzer
from nltk import tokenize
import csv
import datetime
import pandas
from pandas import ExcelFile
from pandas import ExcelWriter
import os

here = os.path.dirname(os.path.abspath(__file__))
subscription_key = "4ae16083b4c1481d9dce8d1cf90b6884"
assert subscription_key

#Read Terms to Search From File
def readSearchTerms():
    terms = []
    filename = os.path.join(here, "searchTerms.xlsx")
    col = pandas.read_excel(filename, sheet_name = 'Sheet1')
    search = col["Search_Term"]
    buckets = col["Bucket_Name"]
    temp = list()
    for i in range(len(col)):
        temp.insert(0, search[i])
        temp.insert(1, buckets[i])
        terms.append(temp[0:2])
    print(terms)
    return terms

#Bing Web Search
def bingWebSearch(term):
    search_url = "https://api.cognitive.microsoft.com/bing/v7.0/news/search"
    headers = {"Ocp-Apim-Subscription-Key" : subscription_key}
    params  = {"q": term, "textDecorations":True, "textFormat":"HTML"}
    response = requests.get(search_url, headers=headers, params=params)
    response.raise_for_status()
    search_results = response.json()
    urls = list()
    for v in search_results["value"]:
        urls.append(v["url"])
    print(urls)
    return urls


def writeUrls(terms):
    file = open("urlsToSearch.txt", "w")
    search = list()
    temp = list()
    for i in terms:
        temp = i
        urls = bingWebSearch(temp[0])
        temp.append(urls)
        search.append(temp)

    #file.close()
    print("Write Complete")
    return search

#Newscrawler is method that walks through indivual articles and gives a sentiment analysis score
#Input url, Output compound sentiment score
def newscrawler(url):
    first = Article(url)
    try:
        first.download()
        print("Downloaded ", first)
        first.parse()
        print("Parsed ", first)
    except:
        print("Bad URL")
        return False
    firstText = first.text
    sid = SentimentIntensityAnalyzer()
    sft = sid.polarity_scores(firstText)
    stri = ''
    for k in sorted(sft):
        temp = sft[k]
        stri = stri + '*' +str(temp)
    myStr = stri.split("*")
    compound = float(myStr[1])
    return compound

#Format for parameters list [['Term', 'Bucket_name', ['url1', 'url2',...]], ...]
def sentAnalysis(search_results):
    analysisList = list()
    for i in search_results:
        urls = i[2]
        compound = list()
        for j in urls:
            compound.append(newscrawler(j))
            total = 0
        for k in compound:
            total = total + k
        if len(compound) != 0:
            avg = total / len(compound)
            now = datetime.datetime.now().date()
            temp = list()
            temp.append(i[0])
            temp.append(i[1])
            temp.append(avg)
            temp.append(now)
            temp.append(urls)
            analysisList.append(temp)

    return analysisList

#Write results to output.csv
def writeToCSV(myList):
    filename = os.path.join(here, "output.csv")
    with open(filename, "a") as f:
        writer = csv.writer(f)
        for i in myList:
            writer.writerow([i[0],i[1],i[2],i[3],i[4]])
    return True

#write time info to log file
def metalog(start, end, total):
    filename = os.path.join(here, "newscrawlerlog.csv")
    now = datetime.datetime.now().date()
    with open(filename, "a") as f:
        writer = csv.writer(f)
        writer.writerow([now, start, end, total])
    return True

#Main
#timekeeping
starttime = datetime.datetime.now()
print(starttime)
#program
terms = readSearchTerms()
terms2 = writeUrls(terms)
result = sentAnalysis(terms2)
print(result)
csv = writeToCSV(result)
#timekeeping
endtime = datetime.datetime.now()
print(endtime)
totaltime = endtime - starttime
print(totaltime)
#metalog
metalog(starttime, endtime, totaltime)