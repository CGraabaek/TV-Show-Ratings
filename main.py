from helpers import *
import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('imdb_series.xlsx')
worksheet = workbook.add_worksheet()
shows = ['2861424','0944947','0141842','0306414','2085059','0773262','7366338','0795176','5491994','0096697','0411008','1856010','0460649','0149460','0386676','0285331','0903747','0185906','4574334','4052886']

worksheet.write(0, 0,  "show")
worksheet.write(0, 1, "rating")
worksheet.write(0, 2, "avg_rating")
worksheet.write(0, 3, "last_episode_rating")

row = 1
col = 0

for show in shows:
    title = get_show_title(show)
    print("Processing show: {} ".format(title))
    rating = get_rating_for_show(show)
    avg,last_episode = get_avg_rating_for_show(show)

    worksheet.write(row, col,     title)
    worksheet.write(row, col + 1, rating)
    worksheet.write(row, col + 2, avg)
    worksheet.write(row, col + 3, last_episode)

    row += 1

workbook.close()
