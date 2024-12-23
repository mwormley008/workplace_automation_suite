from imdb 
import openpyxl

ia = imdb.Cinemagoer()

# from openpyxl import Workbook, load_workbook
'''
alright so this imdb module lets you import the movies you want as dictionary objects,
o you just need to learn which keys you want learn:
plot
director
cast
cover

Ok so we need to search the database with the name of the film then using
that id we can fill in the plot, director, cast, and cover.
'''