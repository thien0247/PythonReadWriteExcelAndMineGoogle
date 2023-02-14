# Python Read and Write MS Excel and Mine Google
## Python Tutorial: How to Read-Write Excel Files, Web-Scrape Google and Create Interactive Maps in 20 lines of Code

The following short article shows just how simple it is to use Python programming language in a data science project. In this example, we’ll first go to Statista.com (public dataset provider) and download the MS Excel dataset that contains the list of 100 largest companies in the world. The file that contains only two columns, the company name and their current market value. Our goal is to use Python to read the rows and cells inside the Excel file and use it to search the internet for some additional information, such as the company’s headquarters location and it’s map coordinates (latitude and longitude). You’ll see how easily this can be done by using Python web-scraping capabilities. We’ll also show how to write the newly found information back into the Excel sheet and use it to create an infographic that shows the headquarter location of 100 of the world’s top companies on the map.

More details at:
https://www.joe0.com/2019/04/20/python-tutorial-reading-writing-excel-files-data-gathering-by-web-scraping-google/

https://www.joe0.com/2019/04/20/python-tutorial-reading-writing-excel-files-data-gathering-by-web-scraping-google/


The following short article shows just how simple it is to use Python programming language in a data science project. In this example, we’ll first go to Statista.com (public dataset provider) and download the MS Excel dataset that contains the list of 100 largest companies in the world. The file contains only two columns, the company name and their current market value. Our goal is to use Python to read the rows and cells inside the Excel file and use it to search the internet for some additional information, such as the company’s headquarters location and it’s map coordinates (latitude and longitude). You’ll see how easily this can be done by using Python web-scraping capabilities. We’ll also show how to write the newly found information back into the Excel sheet and use it to create an infographic that shows the headquarter location of 100 of the world’s top companies on the map.


Let’s start by downloading the dataset which contains the list of Top 100 largest companies in the world by their market value in 2018 (in billion U.S. dollars).

The Excel file can be downloaded from Statista.com, it’s available at the following URL: https://www.statista.com/statistics/263264/top-companies-in-the-world-by-market-value/

The dataset looks like this:



## READING EXCEL FILE
Now, you probably ask, how do we open an Excel file in Python and read the cells in Excel? Well, it can be done in 3 lines of code.
```python
import pandas as pd
df = pd.read_excel("top-companies-in-the-world-by-market-value-2018.xlsx", "Data")
print(df.iloc[0, 0])
```
What does the code above do?

Line 1: Imports the Pandas library that allows us to manipulate Excel files

Line 2: Define ‘df’ which reads the ‘Data’ sheet in the file called: top-companies-in-the-world-by-market-value-2018.xlsx

Line 3: Print the first cell on the first row

What the result looks like, as expected, the first row and first column contains the company name: ‘Apple’



 

## WEB SCRAPING GOOGLE
Now, that we can read the Excel dataset, let’s find the location of headquarters for each of the company names. To do so in a real-life, you’d probably just visit Google and type in “Apple Headquarters” as a query. It would return something like this:



The Google search tells us, that the location of Apple headquarters is in ‘Cupertino, California, United States’.

The same thing can be done in Python. We just need to instruct Python to take the company name from our Excel sheet, do a Google Search for ‘Company Name Headquarters’ and then scrape the name of the city from the source code of the Google result page. Hard? Not really!

Python is well equipped for assignments like these, even though this is not exactly kosher, as Google might not like being scraped. As difficult as it may look, the entire undertaking can be accomplished in just 2 lines of code :)
```python
soup = BeautifulSoup(requests.get('https://www.google.com/search?q=' + df.iloc[0, 0] +'+headquarters', headers=userAgent).text, 'html.parser')
print(soup.find('div', {"class": "Z0LcW"}).get_text())
```
What happened here?

Line 1: Use BeautifulSoup library that allows us to download the Google page of our search and access DOM objects inside the downloaded HTML source code of the search result.

Line 2: Print the DIV with the name ‘Z0LcW’ which contains the location of the company name. Note: The name of the DIV tag can be easily found by doing a search in Google Chrome, inspecting the page, just like this:



Now it’s time to loop over a couple of the company names from our Excel dataset, to see if we can get their headquarters location and coordinates by using Python.

Let’s run the following code:
```python
import pandas as pd, requests as requests; from bs4 import BeautifulSoup
df = pd.read_excel("C:\\0.code\PythonTopCompanies\\top-companies-in-the-world-by-market-value-2018.xlsx", "Data")
userAgent = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36'}
for x in range(0, len(df.index)-95):
    soup = BeautifulSoup(requests.get('https://www.google.com/search?q=headquarters ' + df.iloc[x, 0], headers=userAgent).text, 'html.parser')
    headquarters = soup.find('div', {"class": "Z0LcW"}).get_text()
    soup = BeautifulSoup(requests.get('https://www.google.com/search?q=coordinates ' + headquarters, headers=userAgent).text, 'html.parser')
    coordinates = soup.find('div', {"class": "Z0LcW"}).get_text()
    print(df.iloc[x, 0] + " | " + headquarters + " | " + coordinates)
```
As you can see the code first imports all necessary libraries, then

Reads the Excel dataset
Sets the UserAgent for our browser connection (this is so the Google thinks we’re doing a search from a regular Chrome browser)
Loops over the first 5 company names from the Excel dataset and do a search for the location of the headquarters and its coordinates
Print the results separated by a pipeline character – |
It seems above code works as expected for the first 5 companies:




How simple right?

In just 9 lines of code, we are processing the Excel file line by line, browsing the web, programmatically parsing the Google search result HTML as well as retrieving two totally new pieces of information.

## WRITING EXCEL FILE
First, let’s amend our Excel dataset file, by adding two more columns: Location and Coordinates, these are the new columns where we’ll write our scraped data:



Now save and close the Excel file and let’s amend the Python code so it writes the data to specific above columns. This can be achieved by adding two lines of code into our looping mechanism:
```
df.loc[x, 'Location'] = headquarters
df.loc[x, 'Coordinates'] = coordinates
```

## THE CODE & RESULTS
The entire code to read the Excel file, and scrape info from Google and write the results back to the Excel file looks like this.

You can also download it here: https://github.com/JozefJarosciak/PythonReadWriteExcelAndMineGoogle
```python
import pandas as pd; import requests as requests; import urllib
from bs4 import BeautifulSoup
datasetLocation = "C:\\0.code\PythonTopCompanies\\top-companies-in-the-world-by-market-value-2018.xlsx";
df = pd.read_excel(datasetLocation, "Sheet1")
userAgent = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36'}
searchURL = "https://www.google.com/search?q="
for x in range(0, len(df.index)):
    if str(df.iloc[x, 2]) is None:
        print(str(df.iloc[x, 0]) + " | " + str(df.iloc[x, 2]) + " | " + str(df.iloc[x, 3]))
    else:
        companyName = df.iloc[x, 0]
        search1 = BeautifulSoup(requests.get(searchURL + urllib.parse.quote_plus(companyName) + '+headquarters ', headers=userAgent).text, 'html.parser')
        headquarters = search1.find('div', {"class": "Z0LcW"}).get_text()
        if not headquarters:
            headquarters = search1.find('div', {"class": "desktop-title-subcontent"}).get_text()
        search2 = BeautifulSoup(requests.get(searchURL + urllib.parse.quote_plus(headquarters) + '+coordinates', headers=userAgent).text, 'html.parser')
        coordinates = search2.find('div', {"class": "Z0LcW"}).get_text()
        df.loc[x, 'Location'] = headquarters; df.loc[x, 'Coordinates'] = coordinates
        print(df.iloc[x, 0] + " | " + headquarters + " | " + coordinates)
    df.to_excel(datasetLocation, index=False)
```
And the result is a new Excel file that contains the new Location and Coordinates information:



## MAPPING THE RESULTS IN MICROSOFT BING MAPS
We can do this straight in Excel, where we can insert a Bing Map integration and display the list of TOP 100 firms on the map:



## MAPPING THE RESULTS IN GOOGLE MAPS
Or if you want to take your data further, simply import the entire Excel file into Google to create a fully interactive map. Here is the demo of Top 100 companies by the location of their headquarters (number 1-100), click on the market to see the details:



You can see the full map here: https://www.google.com/maps/d/u/0/viewer?mid=1IbJEik5FwWkUfWkTlQQj8xN_K00zP8T8&ll=26.781517642877535%2C3.937561664639958&z=3

## DATASET DOWNLOAD
For those who wish to download a copy of the dataset we’ve created during this short Python lesson, here is the Excel file:

top-companies-in-the-world-by-market-value-2018

## SCREENSHOT


## CONCLUSION
If all of the above can be done with just 20 lines of code, it’s easy to see why Python is currently the fastest-growing programming language. It is a language of choice especially in the data science field where reading, writing and combining datasets, as well as scraping the web for data is a regular routine. Let me know your comments.
