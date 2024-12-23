import imdb 
import openpyxl

ia = imdb.Cinemagoer()
movies = ia.search_movie('matrix')
# print(movies)
print(movies[0]['title'])
test_movie_id = movies[0].movieID
# print(movies[0].keys())

testmovie = ia.get_movie(test_movie_id)
# print(testmovie)
# trivia_keys = ['stars', 'plot outline', 'plot', 'director', 'synopsis']
trivia_keys = ['stars', 'plot outline', 'director']

for key in trivia_keys:
    if key in testmovie.keys():
        value = testmovie[key]
        if isinstance(value, list):  # Check if it's a list
            # Safely handle mixed types in the list
            names = []
            for person in value:
                if hasattr(person, 'get') and 'name' in person:
                    names.append(person['name'])
                else:
                    names.append(str(person))  # Fallback for non-dict elements
            print(f"{key.capitalize()}: {names}")
        else:  # Handle non-list types
            print(f"{key.capitalize()}: {value}")
    else:
        print(f"{key.capitalize()}: Not available")

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

['title', 'year', 'kind', 'cover url', 'canonical title', 'long imdb title', 'long imdb canonical title', 'smart canonical title', 'smart long imdb canonical title', 'full-size cover url'] 
'''