from __future__ import print_function
from selenium import webdriver
import bs4
import time
from difflib import SequenceMatcher
from datetime import date
import pandas as pd
import sys
from camelcase import CamelCase
import random
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1MqDZLmTBqdpW4boxvjEkDerW7c8mfjFZa7Q69jMapcw'
SAMPLE_RANGE_NAME = 'My Collection'

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

driver = webdriver.Chrome()
#driver.maximize_window()
# =============================================================================
# 
# Excel_Workbook_Name = 'Comics.xlsx'
# Excel_Sheet_Name = 'My Collection'
# =============================================================================

def FetchComics(SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME):
    #Opens the specified google sheet
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    service = build('sheets', 'v4', credentials=creds)
    
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    
    values = result.get('values', [])
    new = pd.DataFrame.from_dict(values)
    new.columns = new.iloc[0]
    
    return(new)


def FindComic():
    return()
    
    
    
# =============================================================================
#     if not values:
#         print('No data found.')
#     else:
#         print('Name, Major:')
#         for row in values:
#             print(row)
#             # Print columns A and E, which correspond to indices 0 and 4.
#             # print('%s, %s' % (row[0], row[3]))
# =============================================================================


Starting_DF = FetchComics(SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME)
sortedsheet = Starting_DF.sort_values(by=['Title','Volume','Issue'])

driver.get("https://comicspriceguide.com/login")

# The error codes
NO_SEARCH_RESULTS_FOUND = 1

# Login Elements
input_login_username = driver.find_element_by_xpath('//input[@id="user_username"]')
input_login_password = driver.find_element_by_xpath('//input[@id="user_password"]')
button_login_submit = driver.find_element_by_id("btnLogin")

# Logging in code
input_login_username.send_keys("botbotbot")
input_login_password.send_keys("iforgotit!")
driver.execute_script("arguments[0].click();",button_login_submit)

# Waiting between 5 and 20 seconds to look like a user
time.sleep(random.uniform(5, 10))

htmlBody = ''
rundate = date.today().strftime("%Y-%m-%d")

#for comic_num in sortedsheet.iterrows():
for index, thisComic in sortedsheet.iterrows():
    try:
        
        title = str(thisComic['Title']).strip().upper()
        issue = str(int(thisComic['Issue'])).strip()
        grade = str(thisComic['Grade']).strip()
        cgc = "No" if thisComic['CGC Graded'] == None else thisComic['CGC Graded']
        variant = '' if str(thisComic['Variant']).strip() == 'nan' else str(thisComic['Variant']).strip()
        url = '' if str(thisComic['Book Link']).strip() == 'nan' else str(thisComic['Book Link']).strip()
        price_paid = "$" + (str(thisComic['Price Paid']) if str(thisComic['Price Paid']) != None else "0")

        fullName = title + " #" + issue + variant
                        
        if url == 'None' :
            print(fullName + " - Searching...")
            # Navigating to search page
            if(driver.current_url != "https://comicspriceguide.com/Search"):
                driver.get("https://comicspriceguide.com/Search")
    
            # Wait for the search page to load.
            driver.implicitly_wait(15)
    
            # The Search elements.
            input_search_title = driver.find_element_by_id("search")
            input_search_issue = driver.find_element_by_id("issueNu")
            button_search_submit = driver.find_element_by_id("btnSearch")
                    
            # Input the search parameters.
            input_search_title.send_keys(title)
            #input_search_title.send_keys(str(thisComic['Title']))
            
            #sleep to prevent overloading the site...
            time.sleep(random.uniform(2, 5))
            input_search_issue.send_keys(issue)
            #input_search_issue.send_keys(int(thisComic['Issue']))
            
            driver.execute_script("arguments[0].click();",button_search_submit)
    
            # Wait for results to show up
            time.sleep(random.uniform(5, 20))
    
            # Source Code of the search result page.
            source_code = driver.page_source
    
            # Instantiate BS4 using the source code.
            soup = bs4.BeautifulSoup(source_code,'html.parser')
    
            # Initial similarity. This similarity is between the given title and the hyperlink comic.
            similarity = 0
    
            # Link of the comic.
            comic_link = ''
            
            percentage = 0
            
            # Checnew = pd.DataFrame.from_dict(data)k all the books on the results screen to determine best match
            for candidate in soup.find_all('a', attrs={'class':'grid_issue'}):
                # Replace the superscript "#" in the comic name
                a = str(candidate.text).replace("<sup>#</sup>","#").upper()
               # Check for similarity between the hyperlink comic and my comic title. If more,
                percentage = similar(a,fullName)
                if percentage > similarity:
                    similarity = similar(a,fullName)
                    final_link = 'https://comicspriceguide.com' + str(candidate["href"])
                    comic_link = final_link
            if percentage > 0 :
                print("     Found a match, confidence: " + str(int(percentage*100)) + "% - " + comic_link) 
        else:
            percentage = None
            print(str(thisComic['Title']) + " #" + str(thisComic['Issue']) + " - " + str(thisComic['Book Link']))
            comic_link = thisComic['Book Link']

# =============================================================================
#  A match has been determined - get the details
# =============================================================================
        if comic_link != '':
            driver.get(comic_link)
        else:
            raise ValueError(NO_SEARCH_RESULTS_FOUND,"Looks like the search gave no result. Try searching the title and issue manually to confirm the issue.",thisComic['Title'],thisComic['Issue'])

        # Wait 5 seconds for page to load and get its source code
        time.sleep(random.uniform(5, 20))
        source_code = driver.page_source

        # New BS4 Instance with the comic's page's source code
        soup = bs4.BeautifulSoup(source_code,'html.parser')

        # Finding out all details
        publisher = soup.find('a',attrs={'id':'hypPub'}).text
        volume = soup.find('span',attrs={'id':'lblVolume'}).text
        notes = soup.find('span',attrs = {'id':'spQComment'}).text
        keyIssue = "Yes" if "Key Issue" in soup.text else "No"
        image = soup.find('img',attrs={'id':'imgCoverMn'})['src']
        basic_info = []
        for s in soup.find_all('div',attrs={"class":"m-0 f-12"}):
            basic_info.append(s.parent.find('span',attrs={"class":"f-11"}).text.replace("   ", " "))
        published = basic_info[0] if basic_info[0] != " ADD" else "Unknown"
        comic_age = basic_info[1] if basic_info[1] != " ADD" else "Unknown"
        cover_price = basic_info[2] if basic_info[2] != " ADD" else "Unknown"   

# =============================================================================
#  Get prices into a dataframe to find our grade
#  Known Defect: Comics graded as 10 fail due to the DF having '10.'not '10.0'
# =============================================================================
        if len(grade) < 3:
            grade = grade + ".0"          
        priceTable = soup.find(name='table',attrs={"id":"pricetable"})
        # Load the priceTable into a dataframe
        pricesdf = pd.read_html(priceTable.prettify())[0]
        # Truncate Condition column values to allow matching
        pricesdf['Condition'] = pricesdf['Condition'].str[:3]
        pricesdf = pricesdf.rename(columns={'Graded Value  *': 'Graded Value'})
        thisbooksgrade = pricesdf.loc[pricesdf['Condition'] == grade]
        RawValue = thisbooksgrade['Raw Value'].iloc[0]
        GradedValue = thisbooksgrade['Graded Value'].iloc[0]
        value = RawValue if cgc.upper() == 'NO' else GradedValue
        
        characters_info = soup.find('div',attrs={'id':'dvCharacterList'}).text if soup.find('div',attrs={'id':'dvCharacterList'}) != None else "No Info Found"
        story = soup.find('div',attrs={'id':'dvStoryList'}).text.replace("Stories may contain spoilers","")
        url_link = driver.current_url
        
# =============================================================================
#  Determine Book Price change from last scan using the "Value" field
# =============================================================================
# =============================================================================
#         LastScanValue = str(thisComic['Value']).strip()
#         if len(LastScanValue) > 0:
#             LastScanValue = float(LastScanValue.replace('$',''))
#             CurrentScanValue = float(value.replace('$',''))
#             priceshift = round((CurrentScanValue - LastScanValue),2)
#             print('PriceShift: ' + str(priceshift))
# =============================================================================
            
# =============================================================================
#  update the DF
# =============================================================================
        sortedsheet.at[index,'Publisher'] = publisher
        sortedsheet.at[index,'Volume'] = volume
        sortedsheet.at[index,'Published'] = published
        sortedsheet.at[index,'KeyIssue'] = keyIssue
        sortedsheet.at[index,'Cover Price'] = cover_price
        sortedsheet.at[index,'Comic Age'] = comic_age
        sortedsheet.at[index,'Notes'] = notes
        sortedsheet.at[index,'Confidence'] = percentage
        sortedsheet.at[index,'Book Link'] = url_link
        sortedsheet.at[index,'Graded'] = GradedValue
        sortedsheet.at[index,'Ungraded'] = RawValue
        sortedsheet.at[index,'Cover Image'] = image
        sortedsheet.at[index, rundate] = value
            
# =============================================================================
#  Data for html layout
# =============================================================================
        if cgc.upper() =='NO':
            cgcdiv = ''
        else: 
            cgcdiv = "<div class='cgc'>CGC</div>"
            
        htmlBody = htmlBody + "<div class='hvrbox'><img src='"  +  str(image) + "' alt='Cover' class='hvrbox-layer_bottom'><div class='hvrbox-layer_top'><div class='hvrbox-text'>" 
        htmlBody = htmlBody + "<a href='" + str(url_link) + "'>" + str(title) + " #" + str(issue) + str(variant) +"<br><br>Grade: " + str(grade) + "<br><br>Value: " + str(value) + "<br><br>" + str(notes) + "</a></div>" + str(cgcdiv) + "</div></div>"


    except ValueError as ve:
        if(ve.args[0] == NO_SEARCH_RESULTS_FOUND):
            print("     Unable to find Match for " + str(ve.args[2]) + " #" + str(ve.args[3]))
        driver.get("https://comicspriceguide.com/Search")
        
    except Exception as e:
        print("Error while working on " + title + ' ' + str(e))
        driver.get("https://comicspriceguide.com/Search")
        continue

# =============================================================================
#  Committed to XLS / Replacing with Google Sheet
# =============================================================================
# with pd.ExcelWriter(Excel_Workbook_Name, mode='w') as writer:  
#     sortedsheet.to_excel(writer, sheet_name=Excel_Sheet_Name)
# =============================================================================
    
    with open("comics.html",'w') as f:
        f.write("""<style type'"text/css">
body {background-color: 282828;}
a {color: whitesmoke;text-decoration: none;}
.cgc {background-color: rgb(148, 7, 35);z-index: inherit 5;font-family: Arial, Helvetica, sans-serif;position: absolute;bottom: 0;}
.hvrbox,
.hvrbox * {box-sizing: border-box; padding: 5px;}
.hvrbox {position: relative;display: inline-block;overflow: hidden;width: 250px;height: 400px;}
.hvrbox img {width: 250px;height: 400px;}
.hvrbox .hvrbox-layer_bottom {display: block;}
.hvrbox .hvrbox-layer_top {
	opacity: 0;
	position: absolute;
	top: 0;
	left: 0;
	right: 0;
	bottom: 0;
	width: 250px;
	height: 400px;
	background: rgba(0, 0, 0, 0.6);
	color: #fff;
	padding: 15px;
	-moz-transition: all 0.4s ease-in-out 0s;
	-webkit-transition: all 0.4s ease-in-out 0s;
	-ms-transition: all 0.4s ease-in-out 0s;
	transition: all 0.4s ease-in-out 0s;
}
.hvrbox:hover .hvrbox-layer_top,
.hvrbox.active .hvrbox-layer_top {opacity: 1;}
.hvrbox .hvrbox-text {
	font-family: Arial, Helvetica, sans-serif;
    text-align: center;
	font-size: 18px;
	display: inline-block;
	position: absolute;
	top: 50%;
	left: 50%;
	-moz-transform: translate(-50%, -50%);
	-webkit-transform: translate(-50%, -50%);
	-ms-transform: translate(-50%, -50%);
	transform: translate(-50%, -50%);
}
.hvrbox .hvrbox-text_mobile {
	font-size: 15px;
	border-top: 1px solid rgb(179, 179, 179); /* for old browsers */
	border-top: 1px solid rgba(179, 179, 179, 0.7);
	margin-top: 5px;
	padding-top: 2px;
	display: none;
}
.hvrbox.active .hvrbox-text_mobile {display: block;}
</style>
""")
        f.write(htmlBody)

print("Work is complete.")
