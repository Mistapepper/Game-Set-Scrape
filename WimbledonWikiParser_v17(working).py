#!/usr/bin/env python
# -*- coding: utf-8 -*-

from bs4 import BeautifulSoup
import requests
import re 
import openpyxl
import os 

#-------------------------------------------------------------------------------
# Following code section altered from https://stackoverflow.com/questions/39442058/scraping-several-pages-with-beautifulsoup

year = 1884 
url = "https://en.wikipedia.org/wiki/{tourneyYear}_Wimbledon_Championships"

wb = openpyxl.load_workbook(os.path.expanduser("~/MyPythonScripts/all_womens_tennis_grand_slam_winners.xlsx"))
sheet = wb.get_sheet_by_name('Sheet1')
winnerCol = sheet['D']
yearCol = sheet['B']
DOFCol = sheet['A']
targetRow = 1

for cell in winnerCol:
    response = requests.get(url.format(tourneyYear=year))
    text = response.text
    soup = BeautifulSoup(text, "html.parser")

    overviewTable = soup.find('table', attrs ={'class':"infobox vevent"})

    if targetRow == 1:
        targetRow += 1
        continue

    if overviewTable is None:
        if cell.row == targetRow:
            cell.value = 'None'
            for cell in yearCol:
                if cell.row == targetRow:
                    cell.value = 'None'
            wb.save('all_womens_tennis_grand_slam_winners.xlsx')
        if year < 2017:
            targetRow += 1
            year += 1
            continue 

    womensSingles = overviewTable.find('a'[0], title=re.compile ("Women's Singles"))

    womensSinglesInfo = womensSingles.parent.parent #this gives you the 'tr' of the 10th 'th' in the 'table'

    womensSinglesChampInfo = womensSinglesInfo.next_sibling.next_sibling

    womensSinglesChampName = womensSinglesChampInfo.find_all('a')[1].text

##    print(str(year) + ': ' + womensSinglesChampName)

    if cell.row == targetRow:
        cell.value = womensSinglesChampName
##        print(cell.value)
        wb.save('all_womens_tennis_grand_slam_winners.xlsx')
            
    for cell in yearCol:
        
        tableHeader = overviewTable.find('th')
        womensSinglesChampYear = tableHeader.text[0:4]

        if cell.row == targetRow:
            cell.value = womensSinglesChampYear
            wb.save('all_womens_tennis_grand_slam_winners.xlsx')

            for cell in DOFCol:
                if cell.row == targetRow:
                    dateHeader = overviewTable.find('th', attrs={"scope":"row"})
                    date = dateHeader.parent
                    nDash = "–"
                    hyphen = "-"
    ##                nDashResult = date.find_next('td', string=re.compile('.*{0}.*'.format(nDash)))
    ##                hyphenResult = date.find_next('td', string=re.compile('.*{0}.*'.format(hyphen)))
                    words = date.find('td').get_text().split()
##                    if year == 1890:
##                        print('boooboooboo')
##                        for content in specialResult:
##                            words = content.split()
##                            print(len(words))
####                        cell.value = 'placeholder'
####                        wb.save('all_womens_tennis_grand_slam_winners.xlsx')
##                        break #if I remove this break, the script returns an error for 1891. Needs fixing.
                    print(words)
##                        print(len(words))
##                        break #if I remove this break, the script returns a non-callable object error for year 1890. Needs fixing.
                    if nDash and hyphen in words:
                        if words.index(nDash) < words.index(hyphen):
                            words.remove(hyphen)
                        else:
                            words.remove(nDash)
                            
                            for index, word in enumerate(words):
                                
                                if word == hyphen:
                                    print(index, word)
                                    if index != 0:
                                        if index != len(words)-1:
        ##                                        print('Whole string: "{0}"'.format(content))
                                            before = words[index-1]
                                            after = words[index+1]
                                            after2 = words[index+2]
                                            if after2 == 'July' and after == 1:
                                                dateDay = 30
                                                dateMonth = 'June'
                                                dateYear = (int(year)-1)
                                                finalDateResult = dateDayStr + ' ' + dateMonth + ' ' + dateYear
                                                cell.value = finalDateResult
                                                print(cell.value + '\n')
                                                wb.save('all_womens_tennis_grand_slam_winners.xlsx')
                                            else: 
                                                try:
                                                    dateDay = (int(after)-1)
                                                except ValueError:
                                                    dateDay = (int(after2)-1)
    ##                                            dateDay = (int(after)-1) 
                                                dateDayStr = str(dateDay)
                                                dateMonth = after2
                                                dateYear = str(year)
                                                finalDateResult = dateDayStr + ' ' + dateMonth + ' ' + dateYear
                                                cell.value = finalDateResult
                                                print(cell.value + '\n')
                                                wb.save('all_womens_tennis_grand_slam_winners.xlsx')
          
                                            
                                else:
                                    if word == nDash:
                                        print(index, word)
        ##                                print('Whole string: "{0}"'.format(content))
                                        if index != 0:
                                            before = words[index-1]
                                            if index != len(words)-1:
        ##                                        print(index)
                                                after = words[index+1]
                                                after2 = words[index+2]
                                                try:
                                                    dateDay = (int(after)-1)
                                                except ValueError:
                                                    dateDay = (int(after2)-1)
                                                dateDayStr = str(dateDay)
                                                dateMonth = after2
                                                dateYear = str(year)
                                                finalDateResult = dateDayStr + ' ' + dateMonth + ' ' + dateYear                          
                                                cell.value = finalDateResult
                                                print(cell.value + '\n')
                                                wb.save('all_womens_tennis_grand_slam_winners.xlsx')

                    else:
                        for index, word in enumerate(words):
                                
                                if word == hyphen:
                                    print(index, word)
                                    if index != 0:
                                        if index != len(words)-1:
        ##                                        print('Whole string: "{0}"'.format(content))
                                            before = words[index-1]
                                            after = words[index+1]
                                            after2 = words[index+2]
                                            try:
                                                dateDay = (int(after)-1)
                                            except ValueError:
                                                dateDay = (int(after2)-1)
                                            dateDay = (int(after)-1)
                                            dateDayStr = str(dateDay)
                                            dateMonth = after2
                                            dateYear = str(year)
                                            finalDateResult = dateDayStr + ' ' + dateMonth + ' ' + dateYear
                                            cell.value = finalDateResult
                                            print(cell.value + '\n')
                                            wb.save('all_womens_tennis_grand_slam_winners.xlsx')
          
                                            
                                else:
                                    if word == nDash:
                                        print(index, word)
        ##                                print('Whole string: "{0}"'.format(content))
                                        if index != 0:
                                            before = words[index-1]
                                            if index != len(words)-1:
        ##                                        print(index)
                                                after = words[index+1]
                                                after2 = words[index+2]
                                                try:
                                                    dateDay = (int(after)-1)
                                                except ValueError:
                                                    dateDay = (int(after2)-1)
                                                dateDayStr = str(dateDay)
                                                dateMonth = after2
                                                dateYear = str(year)
                                                finalDateResult = dateDayStr + ' ' + dateMonth + ' ' + dateYear                          
                                                cell.value = finalDateResult
                                                print(cell.value + '\n')
                                                wb.save('all_womens_tennis_grand_slam_winners.xlsx')



            
##                print(nDashResult)
##                print(hyphenResult)


                
##                for content in nDashResult:
##                    words = content.split()
####                    print(words)
##                    for index, word in enumerate(words):
##                        if word == nDash:
####                            print('Whole string: "{0}"'.format(content))
##                            if index != 0:
##                                before = words[index-1]
##                            if index != len(words)-1:
##                                after = words[index+1]
##                                after2 = words[index+2]
##                            dateDay = (int(after)-1)
##                            dateDayStr = str(dateDay)
##                            dateMonth = after2
##                            dateYear = str(year)
##                            finalDateResult = dateDayStr + ' ' + dateMonth + ' ' + dateYear                          
##                            if cell.row == targetRow:
##                                cell.value = finalDateResult
##                                print(finalDateResult)
##                                wb.save('all_womens_tennis_grand_slam_winners.xlsx')


                
##                else:
##                    for content in hyphenResult:
##                        words = content.split()
##    ##                    print(words)
##                        for index, word in enumerate(words):
##                            if word == hyphen:
##    ##                            print('Whole string: "{0}"'.format(content))
##                                if index != 0:
##                                    before = words[index-1]
##                                if index != len(words)-1:
##                                    after = words[index+1]
##                                    after2 = words[index+2]
##                                dateDay = (int(after)-1)
##                                dateDayStr = str(dateDay)
##                                dateMonth = after2
##                                dateYear = str(year)
##                                finalDateResult = dateDayStr + ' ' + dateMonth + ' ' + dateYear                          
##                                if cell.row == targetRow:
##                                    cell.value = finalDateResult
##                                    print(finalDateResult)
##                                    wb.save('all_womens_tennis_grand_slam_winners.xlsx')
##
##    print(os.getcwd())

    if year is 1891:
##        wb.save('all_womens_tennis_grand_slam_winners.xlsx')
        break  # last page

    year += 1
    targetRow += 1

#-------------------------------------------------------------------------------


