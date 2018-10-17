# Game Set Scrape


This is a Python web scraper. It parses selected Wikipedia pages for info about Wimbledon tennis Grand Slam tournament winners from 1884-2016. It gathers the following data: **Tournament Year, Date of Final, Name of Winner**. It saves the results to an empty .xlsx Excel spreadsheet.

#------------------------------------------------------README-----------------------------------------------------------------------------

## Usage

The script saves the data to an Excel spreadsheet (.xlsx file). You should create an empty spreadsheet and save it to a location on your computer's hard drive. Be sure to update the script to reference the correct file name and location in the `wb` variable:

```python
wb = openpyxl.load_workbook(os.path.expanduser("~/all_womens_tennis_grand_slam_winners.xlsx"))
```

## Requirements 
Beautiful Soup 4 (https://www.crummy.com/software/BeautifulSoup/)  
requests (http://docs.python-requests.org/en/master/)  
re (https://docs.python.org/3/library/re.html)  
os (https://docs.python.org/3/library/os.html)  
openpyxl (https://openpyxl.readthedocs.io/en/stable/)  

