from bs4 import BeautifulSoup
import requests
import openpyxl
import re

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Movies List"
sheet.append(['Rank','Movie Name','Year of Release' , 'Rating'])


try:
    url      = "https://www.imdb.com/chart/top/"
    response = requests.get(url)                                # if good link response 200
    soup     = BeautifulSoup(response.text,'html.parser')       # html code extracted
    movies   = soup.find('tbody',class_='lister-list').find_all('tr')
    
    for movie in movies:
        #print(movie)
        movie_rank = movie.find('td',class_='titleColumn').text.split('.')[0]
        movie_rank = movie_rank.strip()

        movie_name = movie.find('td',class_='titleColumn').a.text

        movie_rate = movie.find('td',class_='ratingColumn imdbRating').strong.text

        movie_year = movie.find('td',class_='titleColumn').span.text         #(1997)
        ''' Using regex method to remove Parantheis'''
        movie_year = re.sub(r'[()]', '', movie_year)

        #print(movie_rank,movie_name,movie_year,movie_rate)
        #print('******'*5)

        sheet.append([movie_rank,movie_name,movie_year,movie_rate])
    

except Exception as msg:
    print('Error msg : ',msg)                 
    

excel.save('Mymovies.xlsx')
print('New Excel File created ')