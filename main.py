import openpyxl
import time
import requests
import urllib
from urllib.request import Request
from bs4 import BeautifulSoup
import tweepy

# Takes the path of the excel file and the sheetname
# For each row in the sheet - if not already processed then searches for Moneycontrol link
# and gets website link and email ID and writes to excel file
def getWebsiteLinks(excelPath, sheetname):
    workbook = openpyxl.load_workbook(excelPath)
    worksheet = workbook[sheetname]

    # Iterate through all the rows and process each row if not already processed
    rowCount = 0
    for row in worksheet:
        companyName = row[0].value
        processingStatus = row[9].value

        if companyName == None: # reached the end of the file, save workbook and stop processing
            workbook.save(excelPath)
            break
        elif processingStatus == None:
            print(companyName)
             # Search company name on Moneycontrol website using DuckDuckGo and get the first link
            searchHeader = {
                'User-Agent':
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36 Edge/18.19582"
            }
            baseSearchString = "site:www.moneycontrol.com/india/stockpricequote" # "https://duckduckgo.com/html/?q=
            searchName = companyName.split(" Ltd.")[0]
            finalSearchString = searchName + ' ' + baseSearchString #'"' + searchName + '" ' + baseSearchString
            baseMoneycontrolLink = "www.moneycontrol.com/india/stockpricequote"

            try:
                linkMoneycontrol = searchOnDuckDuckGo(companyName, baseMoneycontrolLink, searchHeader)
                print(linkMoneycontrol)
                time.sleep(30)  # Sleep for 30 seconds after processing 1 row

                linkWebsite = getWebsiteFromMoneycontrol(linkMoneycontrol)
                emailID = getEmailFromMoneycontrol(linkMoneycontrol)
                twitterHandle = getTwitterHandleFromHomepage(linkWebsite)
                twitterLink = ""
                if twitterHandle != "":
                    twitterLink = "https://twitter.com/" + twitterHandle

                row[4].value = linkMoneycontrol
                row[5].value = linkWebsite
                row[6].value = emailID
                row[7].value = twitterLink
                row[8].value = twitterHandle
                row[9].value = "Done"
                workbook.save(excelPath)
            except Exception as err:
                print(err)
                workbook.save(excelPath)
                break # If any error then move to next row

        rowCount = rowCount + 1
        #if rowCount == 100:
        #    workbook.save(excelPath)
         #   break

def searchOnDuckDuckGo(companyName, baseMoneycontrolLink, searchHeader):
    baseSearchString = "http://duckduckgo.com/html/?q="
    searchInMoneyControl = "site:" + baseMoneycontrolLink
    searchCompanyName = urllib.parse.quote(companyName) #'"'+urllib.parse.quote(companyName)+'"'
    finalSearchString = baseSearchString + searchCompanyName + '%20' + searchInMoneyControl
    print("DuckDuckGo Search String:", finalSearchString)

    try:
        req = Request(finalSearchString,headers=searchHeader)
        site = urllib.request.urlopen(req)
        data = site.read()
        parsed = BeautifulSoup(data,'html.parser')
        data = parsed.find(attrs={'class' : 'result__url'})
        first_link = data.text.strip()
        first_link = 'http://' + first_link
        return first_link
    except Exception as err:
        print(err)
        return ""

def getWebsiteFromMoneycontrol(linkMoneycontrol):
    header = {
        'User-Agent':
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36 Edge/18.19582"
    }
    try:
        html = requests.get(linkMoneycontrol, headers=header).text
        linkWebsite = html.split('<span>Internet</span>')[1]
        linkWebsite = linkWebsite.split('<p><a href="')[1]
        linkWebsite = linkWebsite.split('"')[0]
    except Exception as err:
        print(err)
        return ""
    return linkWebsite

def getEmailFromMoneycontrol(linkMoneycontrol):
    header = {
        'User-Agent':
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36 Edge/18.19582"
    }
    try:
        html = requests.get(linkMoneycontrol, headers=header).text
        emailID = html.split('<span>Email</span>')[1]
        emailID = emailID.split('mailto:')[1]
        emailID = emailID.split('"')[0]
    except Exception as err:
        print(err)
        return ""
    return emailID

def getTwitterHandleFromHomepage(linkWebsite):
    header = {
        'User-Agent':
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36 Edge/18.19582"
    }
    try:
        html = requests.get(linkWebsite, headers=header).text

        twitterHandle = html.split("https://twitter.com/")[1]
        twitterHandle = twitterHandle.split('"')[0]
    except Exception as err:
        print(err)
        return ""

    return twitterHandle

def getTwitterUserInfoForAll(excelPath, sheetName):
    workbook = openpyxl.load_workbook(excelPath)
    worksheet = workbook[sheetName]

    # Iterate through all the rows and clean the twitter handles if not already cleaned
    rowCount = 0
    for row in worksheet:
        companyName = row[0].value
        twitterHandle = row[8].value
        twitterCleaned = row[10].value
        cleanTwitterHandle = ""
        cleanTwitterLink = ""

        if companyName == None: # reached the end of the file, save workbook and stop processing
            workbook.save(excelPath)
            break
        elif twitterCleaned == None:
            # String processing of the Twitter Handle text from excel file to remove unwanted / junk portions
            # Split by , ? / ' space
            try:
                print("Original: " + twitterHandle)
                cleanTwitterHandle = twitterHandle.split(",")[0]
                cleanTwitterHandle = cleanTwitterHandle.split("?")[0]
                cleanTwitterHandle = cleanTwitterHandle.split("/")[0]
                cleanTwitterHandle = cleanTwitterHandle.split(" ")[0]
                cleanTwitterHandle = cleanTwitterHandle.split("'")[0]
            except:
                cleanTwitterHandle = ""

            print("Clean: " + cleanTwitterHandle)

            # Set the cleanTwitterHandle, Twitter Link and Processing Status
            if cleanTwitterHandle != "":
                cleanTwitterLink = "https://twitter.com/" + cleanTwitterHandle

            row[11].value = cleanTwitterHandle
            row[12].value = cleanTwitterLink
            row[10].value = "Done"
    # We have cleaned up all the rows - save the excel file with our edits
    workbook.save(excelPath)

    # Now let us get info from Twitter handle using Twitter API
    for row in worksheet:
        companyName = row[0].value
        cleanTwitterHandle = row[11].value
        twitterInfoProcessed = row[16].value

        if companyName == None: # reached the end of the file, save workbook and stop processing
            workbook.save(excelPath)
            break
        elif cleanTwitterHandle == None:
            # If we don't have twitter handle then mark as processed and move to next
            row[16].value = "Done"
        elif twitterInfoProcessed == None:
            # Use Twitter API to get info
            twitterFollowersCount = None
            twitterPostsCount = None
            twitterCreateDate = None

            twitterFollowersCount, twitterPostsCount, twitterCreateDate = getTwitterUserInfoFromHandle(cleanTwitterHandle)
            time.sleep(0.5)  # Sleep for 0.5 second after processing 1 row
            row[13].value = twitterFollowersCount
            row[14].value = twitterPostsCount
            row[15].value = twitterCreateDate
            row[16].value = "Done"
    # We have cleaned up all the rows - save the excel file with our edits
    workbook.save(excelPath)

# Function to get Twitter Info from a handle using Twitter API
def getTwitterUserInfoFromHandle(twitterHandle):
    # assign the values accordingly
    twitterAPI_consumer_key = "zDEc57LPywYffyN3xrvVJRXCi"
    twitterAPI_consumer_secret = "6vY5hyclXXgwdCgedHiXKOWHqEHatyRLRCBxp99vq2J8RWQTjs"
    twitterAPI_access_token = "4321094953-AY6gVqy41tgUHpkDPGdCL3pAMVAUFwyaRWc7xlZ"
    twitterAPI_access_token_secret = "CBD7I7vSzUz1yEEzAYHV7qP5CzlWnobrzvoeJGBUPHg8M"

    auth = tweepy.OAuth1UserHandler(
        twitterAPI_consumer_key, twitterAPI_consumer_secret, twitterAPI_access_token, twitterAPI_access_token_secret
    )
    followers_count = None
    posts_count = None
    create_date = None
    try:
        # calling the api
        api = tweepy.API(auth)
        # fetching the user
        user = api.get_user(screen_name = twitterHandle)
        # fetching the followers_count
        followers_count = user.followers_count
        posts_count = user.statuses_count
        create_date = user.created_at
        # Convert the date to string and extract the date portion - we don't need the timestamp
        create_date = str(create_date).split(" ")[0]

        print("Handle: " + twitterHandle)
        print("Followers: " + str(followers_count))
        print("Create Date: " + str(create_date))
    except:
        # Do nothing - just return the current values
        print(Exception)

    return followers_count, posts_count, create_date

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    excelPath = 'C://Users//Varun//Documents//ISB AMPBA//05. Data Collection//09. Group Assignment//Indian Stocks Dump.xlsx'
    sheetName = 'Twitter Info'
    # Use this function to get Moneycontrol website and additional info based on company name
    # getWebsiteLinks(excelPath, sheetName)

    # Use this function to get Twitter info for all the companies where Twitter handle is present
    getTwitterUserInfoForAll(excelPath, sheetName)