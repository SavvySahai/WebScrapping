from bs4 import BeautifulSoup #installed and imported beautifulsoup4 and requests
import requests, openpyxl
excel = openpyxl.Workbook()
print(excel.sheetnames)  #Created excel sheet
sheet = excel.active
sheet.title = 'Top Rated Movies' #changed the sheet name

sheet.append(['Movie rank','Movie Name', 'Year of Release', 'IMDB Rating'])
print(excel.sheetnames)

try:
  source = requests.get('https://www.imdb.com/chart/top/')
  source.raise_for_status() #checking if the given link is opening in webpage

  soup = BeautifulSoup(source.text,'html.parser') 
  movies = soup.find('tbody', class_="lister-list").find_all('tr')
  for movie in movies:
      name = movie.find('td', class_="titleColumn").a.text
      rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0] #now rank will be split using . rank will be in index 0 and the rest in index 1
      year = movie.find('td', class_="titleColumn").span.text.strip('()')
      rating = movie.find('td', class_="ratingColumn imdbRating").strong.text #now extracted all the details required for the firstt movie because of break

      
      
      print(rank, name, year, rating)
      sheet.append([rank, name, year, rating]) #this will load each loop into the excel
    

except Exception as e:
    print(e)
excel.save('IMDB Movie Ratings.xlsx') #Saving the excel file