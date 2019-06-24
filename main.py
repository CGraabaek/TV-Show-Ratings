from imdb import IMDb
import xlsxwriter

# create an instance of the IMDb class
ia = IMDb()
#series = ia.get_movie('0944947') #Game of thrones
#series = ia.get_movie('5715524') #Mayans MC 

series = ia.get_movie('0773262')  #Dexter
ia.update(series, 'episodes')
title = series["title"]

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook(title+'.xlsx')
worksheet = workbook.add_worksheet()

# Start from the first cell. Rows and columns are zero indexed.
worksheet.write(0, 0,  "season")
worksheet.write(0, 1, "episode")
worksheet.write(0, 2, "title")
worksheet.write(0, 3, "rating")
worksheet.write(0, 4, "votes")

row = 1
col = 0

for season_nr in sorted(series['episodes']):
    for episode_nr in sorted(series['episodes'][season_nr]):
        episode = series['episodes'][season_nr][episode_nr]
       
        worksheet.write(row, col,     season_nr)
        worksheet.write(row, col + 1, episode_nr)
        worksheet.write(row, col + 2, episode["title"])
        worksheet.write(row, col + 3, episode.get('rating'))
        worksheet.write(row, col + 4, episode.get('votes'))
        row += 1

        print('episode #%s.%s; title: %s; rating: %s; votes: %s' %
              (season_nr, episode_nr,episode["title"] ,episode.get('rating'), episode.get('votes')))

workbook.close()