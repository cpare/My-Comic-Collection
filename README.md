# My-Comic-Collection
  Python script that takes just a few inputs from an Excel sheet to query for additional details from https://comicspriceguide.com.  
Result is an updated HTML page and XLS with current values, description, and other helpful features

<b>Provide the following items in an XLS:</b>
 - Title - Your Comics Title (Amazing Spider Man)
 - Issue - The Issue Number (101)
 - Grade - Grade of the book based on numbers offered on https://comicspriceguide.com
 - CGC - Yes/No if the Book has been "Officially graded"
 - Variant - Some Issues have multiple printes/covers, this is where yuo indicate the specific cover you have (E)
 - Price_Paid - The Proce you paid for the book
 - Book Link - Sometimes the app doesn't select the right book, this allows you to specifically provide the URL, when populated we skip the search and use this (Optional)
 
 <b>XLS Results:</b>
  - Publisher - Publisher of the title
  - Title - (Provided Above)
  - Volume - Volume 
  - Issue - (Provided Above)
  - Variant - (Provided Above)
  - Grade - (Provided Above)
  - CGC - (Provided Above)
  - PublishDate - Original Publish Date
  - KeyIssue - Is the book considered a "Key Issue" (Yes/No)
  - Price_Paid - (Provided Above)
  - Cover_Price - Original Cover Price
  - Graded Price - Value of book & and grade if professionally graded
  - Ungraded Price - Value of book if not professionally graded
  - Value - Value of book in current state
  - Comic_Age - Comic "Age" (Silver, Modern)
  - Notes - Short description of the comic, including references to 1st appearances
  - Confidence - Confidence the app had when attempting to match
  - Book Link - (Provided Above)
  - Cover Url - Cover Image
  - Date - The date the scan was performed and the value of the book at that time, allowing you to scan periodically and easily identify price fluctuations.

<b>HTML Results:</b>
 - A single HTML page with the covers of each of your books, formatted to look stunning on any format
 - Hover-over any book to see book details (Title, Grade, Value, CGC Graded)
 - Click any book to go directly to that book on https://comicspriceguide.com

IMPORTANT NOTE: Though I have taken an effort to try and simulate a normal user session wtih delays there is a chance the bot could be detected.  If you run this and are detected by https://comicspriceguide.com you stand a chance of having your IP banned from that site.
