from selenium import webdriver
import bs4
import time
from difflib import SequenceMatcher
from datetime import date
import pandas as pd
import sys


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

driver = webdriver.Chrome()
#driver.maximize_window()

ExcelWorkbookName = 'Comics.xlsx'

rawsheet = pd.read_excel(open(ExcelWorkbookName, 'rb'), sheet_name='My Collection')
sortedsheet = rawsheet.sort_values(by=['Title','Volume','Issue #'])
del rawsheet

# Create an empty Dataframe for our results
dfResults = pd.DataFrame(columns = ['title','issue','grade','cgc','publisher',
                                    'volume','published','keyIssue','price_paid',
                                    'cover_price','Graded','Ungraded','value','comic_age','notes','confidence',
                                    'characters_info','story','url_link'])

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

# Waiting time for page to load.
time.sleep(5)

htmlBody = ''

#for comic_num in sortedsheet.iterrows():
for index, thisComic in sortedsheet.iterrows():
    try:
        title = str(thisComic['Title']).strip().upper()
        issue = str(int(thisComic['Issue #'])).strip()
        grade = str(int(thisComic['Grade'])).strip()
        cgc = "No" if thisComic['CGC Graded'] == None else thisComic['CGC Graded']
        variant = '' if str(thisComic['Variant']).strip() == 'nan' else str(thisComic['Variant']).strip()
        url = '' if str(thisComic['Link']).strip() == 'nan' else str(thisComic['Link']).strip()
        price_paid = "$" + (str(thisComic['Price Paid']) if str(thisComic['Price Paid']) != None else "0")

        fullName = title + " #" + issue + variant
        
        if url == '' :
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
                    
            # Appending the title to the end of list. Example, "The Amazing Spider-Man #101"
            # Not sure why we do this yet
            #FullName = Title + " #" + Issue + Variant
            #comic.append(comic[1].upper() + " #" + str(comic[3]))
    
# =============================================================================
#             # This is to convert "101.0" to just "101"
#             if isinstance(comic[3], float):
#                 comic[3] = int(comic[3])
# =============================================================================
    
            # Input the search parameters.
            input_search_title.send_keys(str(thisComic['Title']))
            time.sleep(1)
            input_search_issue.send_keys(int(thisComic['Issue #']))

            #time.sleep(1)
            driver.execute_script("arguments[0].click();",button_search_submit)
    
            # Wait for results to show up
            time.sleep(6)
    
            # Source Code of the search result page.
            source_code = driver.page_source
    
            # Instantiate BS4 using the source code.
            soup = bs4.BeautifulSoup(source_code,'html.parser')
    
            # Initial similarity. This similarity is between the given title and the hyperlink comic.
            similarity = 0
    
            # Link of the comic.
            comic_link = ''
            
            percentage = 0
            
            # Check all the books on the results screen to determine best match
            for link in soup.find_all('a', attrs={'class':'grid_issue'}):
                # Replace the superscript "#" in the comic name
                a = str(link.text).replace("<sup>#</sup>","#").upper()
               # Check for similarity between the hyperlink comic and my comic title. If more,
                percentage = similar(a,fullName)
                if percentage > similarity:
                    similarity = similar(a,fullName)
                    final_link = 'https://comicspriceguide.com' + str(link["href"])
                    comic_link = final_link
            if percentage > 0 :
                print("     Found a match, confidence: " + str(int(percentage*100)) + "%")
        else:
            percentage = 1
            print(str(thisComic['Title']) + " #" + str(thisComic['Issue #']) + " - " + str(thisComic['Link']))
            comic_link = thisComic['Link']

        # Goto the comic page if result is found.
        if comic_link != '':
            driver.get(comic_link)
        else:
            raise ValueError(NO_SEARCH_RESULTS_FOUND,"Looks like the search gave no result. Try searching the title and issue manually to confirm the issue.",thisComic['Title'],thisComic['Issue #'])

        # Wait 5 seconds for page to load and get its source code
        time.sleep(2)
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
# =============================================================================
        priceTable = soup.find(name='table',attrs={"id":"pricetable"})
        # Load the priceTable into a dataframe
        pricesdf = pd.read_html(priceTable.prettify())[0]
        # Clean up the Condition column (has white spaces)
        pricesdf['Condition'] = pricesdf['Condition'].str[:3]
        pricesdf = pricesdf.rename(columns={'Graded Value  *': 'Graded Value'})
        thisbooksgrade = pricesdf.loc[pricesdf['Condition'] == '9.0']
        RawValue = thisbooksgrade['Raw Value'].iloc[0]
        GradedValue = thisbooksgrade['Graded Value'].iloc[0]
        value = RawValue if cgc == 'No' else GradedValue
        
        characters_info = soup.find('div',attrs={'id':'dvCharacterList'}).text if soup.find('div',attrs={'id':'dvCharacterList'}) != None else "No Info Found"
        story = soup.find('div',attrs={'id':'dvStoryList'}).text.replace("Stories may contain spoilers","")
        url_link = driver.current_url
        
# =============================================================================
#  Data to be put into excel file
# =============================================================================
        dfResults = dfResults.append({'title' : title,
                                      'issue' : issue,
                                      'grade':grade,
                                      'cgc':cgc,
                                      'publisher':publisher,
                                      'volume':volume,
                                      'published':published,
                                      'keyIssue':keyIssue,
                                      'price_paid':price_paid,
                                      'cover_price':cover_price,
                                      'value':value,
                                      'comic_age':comic_age,
                                      'notes':notes,
                                      'confidence':percentage,
                                      'characters_info':characters_info,
                                      'story':story,
                                      'url_link':url_link,
                                      'Graded':GradedValue,
                                      'Ungraded':RawValue
                                      },
                                      ignore_index=True)
        
# =============================================================================
#  Data for html layout
# =============================================================================
        htmlBody = htmlBody + "<div class='hvrbox'><img src='"  +  str(image) + "' alt='Cover' class='hvrbox-layer_bottom'><div class='hvrbox-layer_top'><div class='hvrbox-text'>" 
        htmlBody = htmlBody + "<a href='" + str(url_link) + "'>" + str(title) + " #" + str(issue) +"<br><br>Grade: " + str(grade) + "<br><br>Value: " + str(value) + "<br><br>" + str(notes) + "</a></div></div></div>"


    except ValueError as ve:
        if(ve.args[0] == NO_SEARCH_RESULTS_FOUND):
            print("     Unable to find Match for " + str(ve.args[2]) + " #" + str(ve.args[3]))
            dfResults = dfResults.append({'title' : title,
                                          'issue' : issue,
                                          'grade' : grade,
                                          'cgc' : cgc}, ignore_index=True)
        driver.get("https://comicspriceguide.com/Search")
        
    except Exception as e:
        print("Error: " + str(e))
        dfResults = dfResults.append({'title' : title,
                                      'issue' : issue,
                                      'grade': grade,
                                      'cgc': cgc}, ignore_index=True)
        driver.get("https://comicspriceguide.com/Search")
        continue

   
sheetname = date.today().strftime("%Y-%m-%d")
with pd.ExcelWriter(ExcelWorkbookName, mode='a') as writer:  
    dfResults.to_excel(writer, sheet_name=sheetname)
    
    with open("comics.html",'a') as f:
        f.write("""<style type'"text/css">
body {background-color: 282828;}
a {color: whitesmoke;text-decoration: none;}
.hvrbox,
.hvrbox * {
	box-sizing: border-box;
    padding: 5px;
}
.hvrbox {
	position: relative;
	display: inline-block;
	overflow: hidden;
	width: 250px;
	height: 400px;
}
.hvrbox img {
	width: 250px;
    height: 400px;
}
.hvrbox .hvrbox-layer_bottom {
	display: block;
}
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
.hvrbox.active .hvrbox-layer_top {
	opacity: 1;
}
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
.hvrbox.active .hvrbox-text_mobile {
	display: block;
}
                </style>""")
        f.write(htmlBody)

print("Work is complete.")
