import xlrd
import tkinter
from tkinter import *
from tkinter import ttk, Label, Entry, StringVar
import urllib.request
import urllib.parse
from pytube import YouTube
# from os import startfile


def make_data():
    workbook = xlrd.open_workbook('Recipe.xls')
    worksheet = workbook.sheet_by_index(0)
    global rowc
    rowc = 0
    num_cols = worksheet.ncols  # Number of columns
    global data
    data = []
    for row_idx in range(0, worksheet.nrows):  # Iterate through rows
        data.append([])
        rowc = rowc + 1
        for col_idx in range(0, num_cols):  # Iterate through columns
            cell_obj = worksheet.cell(row_idx, col_idx)  # Get cell object by row, col
            temp = cell_obj.value
            if cell_obj.value != "":
                data[row_idx].append(temp)


def find_synonyms(s):
    workbook = xlrd.open_workbook('Synonyms.xls')
    worksheet = workbook.sheet_by_index(0)
    row = worksheet.nrows
    col = worksheet.ncols

    for x in range(0, row):
        for y in range(0, col):
            if s == worksheet.cell(x, y).value:
                s = worksheet.cell(x, 0).value

    return s


def compute_adjacent_ingredients():
    global user
    user = entry.get()

    global n
    n = int(entry2.get())

    global all_current_matches
    all_current_matches = []

    if user == "salt" or user == "garlic" or user == "water" or user == "onions":
        a = user + " is too common of an ingredient"
        Label(text=a, bg="#DDFFDD", font="Arial 15 bold").pack(side=TOP)
        return

    user = find_synonyms(user)
    global adjacent
    adjacent = []

    for x in range(0, rowc):
        colc = len(data[x])
        for y in range(0, colc):
            if data[x][y] == user:
                for z in range(0, colc):
                    if data[x][z] != user:
                        adjacent.append(data[x][z])

    if len(adjacent) == 0:
        a = "Your ingredient is not in the dataset"
        Label(text=a, bg="#DDFFDD", font="Arial 15 bold").pack(side=TOP)
        return

    find_most_popular()


def find_most_popular():
    row = len(adjacent)
    ingredient_indexes = []
    row = max(row, 5)

    all_current_matches.append(user)

    for r in range(0, row):
        current = adjacent[r]
        counter = 0
        for c in range(0, row):
            if adjacent[c] == current and r != c:
                counter = counter + 1

        ingredient_indexes.append(counter)

    for z in range(0, n):
        max_ingre_index = ingredient_indexes.index(max(ingredient_indexes))
        temp = ingredient_indexes[max_ingre_index]
        if temp == 0:
            continue
        a = str(adjacent[max_ingre_index]) + " " + str(temp) + " occurances"

        # used to find the recipe
        all_current_matches.append(adjacent[max_ingre_index])

        Label(text=a, bg="#DDFFDD", font="Arial 15 bold").pack(side=TOP)
        for i in range(0, len(ingredient_indexes)):
            if ingredient_indexes[i] == temp:
                ingredient_indexes[i] = 0
        find_alternatives(adjacent[max_ingre_index])


def find_alternatives(s):
    workbook = xlrd.open_workbook('Recipes2.xls')
    worksheet = workbook.sheet_by_index(0)
    data2 = []
    for ri in range(0, worksheet.nrows):  # Iterate through rows
        data2.append([])
        for ci in range(0, worksheet.ncols):  # Iterate through columns
            cell_obj = worksheet.cell(ri, ci)  # Get cell object by row, col
            if cell_obj.value != "":
                data2[ri].append(cell_obj.value.lower())

    for x in range(0, len(data2)):
        cc = len(data2[x])
        if s == data2[x][0]:
            for y in range(1, cc):
                a = "one alternative for " + s + ":" + data2[x][y]
                Label(text=a, bg="#DDFFDD", font="Arial 15 bold").pack(side=TOP)


def load_recipes():
    workbook = xlrd.open_workbook('new_epi_r.xls')
    worksheet = workbook.sheet_by_index(0)
    row = worksheet.nrows
    col = worksheet.ncols
    global recipes
    recipes = []

    for x in range(0, row):
        recipes.append([])
        for y in range(0, col):
            recipes[x].append(worksheet.cell(x, y).value)


# all_current_matches
def find_recipe():

    recipe_index = []
    if len(all_current_matches) == 0:
        return

    for i in range(0, len(recipes[0])):
        for j in range(0, len(all_current_matches)):
            if all_current_matches[j] == recipes[0][i]:
                recipe_index.append(i)

    max_fit = 0
    best_fit_index = 0
    current_fit = 0
    for i in range(1, len(recipes)):
        for j in range(0, len(recipe_index)):
            if int(recipes[i][recipe_index[j]]) == 1:
                current_fit = current_fit + 1
        if current_fit > max_fit:
            max_fit = current_fit
            best_fit_index = i
        current_fit = 0

    a = recipes[best_fit_index][0]
    print(a, max_fit)
    search = "how to make " + a + " " + all_current_matches[0]
    query_string = urllib.parse.urlencode({"search_query": search})
    html_content = urllib.request.urlopen("http://www.youtube.com/results?" + query_string)
    search_results = re.findall(r'href=\"\/watch\?v=(.{11})', html_content.read().decode())
    url = "http://www.youtube.com/watch?v=" + search_results[0]
    YouTube(url).streams.first().download(filename='video')
   # try:
     #   startfile("video.webm")
    #except:
     #   try:
      #      startfile("video.mp4")
      #  except:
       #     print("cant play video")


def run_program():
    # create the window
    root = Tk()
    root.title("Recipe Lookup")

    root.configure(background="#DDFFDD")
    photo = tkinter.PhotoImage(file="food.gif")
    window = Label(text="Enter your desired ingredient", bg="#DDFFDD", font="Arial 28 bold").pack(side=TOP, padx=10, pady=10)
    tkinter.Label(window, image=photo).pack(side=TOP)

    make_data()

    global entry
    label1 = StringVar()
    label1.set("Ingredient")
    label1dir = Label(root, textvariable=label1, bg="#DDFFDD", height=4,font="Arial 15 bold")
    label1dir.pack(side="left", padx=10, pady=0)

    entry = Entry(root, width=30)
    entry.pack(side="left", padx=10, pady=10)

    global entry2
    label2 = StringVar()
    label2.set("Num of results")
    label2dir = Label(root, textvariable=label2, bg="#DDFFDD", height=4, font="Arial 15 bold")
    label2dir.pack(side="left", padx=10, pady=10)

    entry2 = Entry(root, width=30)
    entry2.pack(side="left", padx=10, pady=10)

    # ttk.Button(root, text="Exit", command=root.quit).pack(side=BOTTOM, pady=0)
    ttk.Button(root, text="Search", command=compute_adjacent_ingredients).pack(side=BOTTOM, pady=0, expand=4)

    root.mainloop()


if __name__ == '__main__':
    run_program()
    load_recipes()
    find_recipe()
