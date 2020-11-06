from tkinter import *
import wget
import requests
import xlsxwriter

# Variables containing API information which is used in making request and fetching data.
main_url = 'https://api.themoviedb.org/3/search/movie?'
api = 'api_key=1dcf69b9b95240032c80e5d374ca2bee'
search_query = '&query='
no_of_pages = '&page='

# Variables to store search terms, obtained search results and iterators to traverse through search results.
movie_name = None
movie_data = None
total_pages = None
total_results = None
index = None
currPage = None

# Variables used for storing data in Excel
lines_XLSX = 2
Movie_Name = []
Movie_Overview = []


# Sends request to fetch first set of results in JSON format for given search term and calls DataManipulation() function
def MovieSearch():
    global index, searchTerm, movie_name, movie_data, currPage, total_pages, total_results
    parsedTerm = searchTerm.get()
    index = 0
    currPage = 0
    if parsedTerm == '':
        exit()
    movie_name = parsedTerm.replace(' ', '+')
    requested_data = requests.get(url=main_url+api+search_query+movie_name)
    movie_data = requested_data.json()
    total_pages = int(movie_data['total_pages'])
    total_results = int(movie_data['total_results'])
    if (total_pages != 0):
        currPage = 1
    MoviesData.insert(END, f'\n Search Results for \"{parsedTerm}\"')
    MoviesData.insert(END, f'\n Page \"{currPage}\" of \"{total_pages}\"')
    DataManipualation()


# Determines which set of procedures to execute based on number of results and calls ShowData() function
def DataManipualation():
    global currPage, total_pages, total_results
    if total_pages == 0:
        print('No Results Found')
        MoviesData.insert(END, '\n No Results Found')
    elif total_pages == 1:
        currResults = total_results
        ShowData(currResults)
    elif (total_pages > 1):
        ShowData(20)
        moreBtn.grid()


# Sends request to fetch more results from API and store in JSON format
def RequestData(currPage):
    global movie_name, movie_data, total_pages, total_results
    requested_data = requests.get(
        url=main_url+api+search_query+movie_name+no_of_pages+str(currPage))
    movie_data = requested_data.json()


# Extracts data by parsing JSON and feeds to Text Widget using tkinter
def ShowData(currResults):
    global movie_data, index
    for i in range(currResults):
        index += 1
        movieName = movie_data['results'][i]['original_title']
        overview = movie_data['results'][i]['overview']

        print('\n\n Name Of Movie ', index, ': '+movieName)
        MoviesData.insert(END, f'\n\nName Of Movie {index}: {movieName}')
        Movie_Name.append(movieName)

        if(overview != ''):
            print(' Overview Of Movie: '+overview, '\n')
            MoviesData.insert(END, f'\nOverview Of Movie: {overview}')
            Movie_Overview.append(overview)
        else:
            print(' Overview Of Movie: NA')
            Movie_Overview.append('NA')
            MoviesData.insert(END, '\n Overview Of Movie: NA')
    SaveData(currResults)


# Saves data in Excel format with file name as movie.xlsx
def SaveData(currResults):
    global lines_XLSX
    workbook = xlsxwriter.Workbook('movie.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Movie Name')
    worksheet.write('B1', 'Movie Description')
    for i in range(2, lines_XLSX+currResults):
        wb = 'A'+str(i)
        wbb = 'B'+str(i)
        worksheet.write(wb, Movie_Name[i-2])
        worksheet.write(wbb, Movie_Overview[i-2])
    lines_XLSX += currResults
    workbook.close()


# Loads next set of results when More button is clicked
def loadMore():
    global currPage, total_pages, total_results
    currPage += 1
    MoviesData.config(state="normal")
    MoviesData.delete('3.0', END)
    MoviesData.insert(END, f'\n Page \"{currPage}\" of \"{total_pages}\"')
    
    if currPage < total_pages:
        RequestData(currPage)
        ShowData(20)
    elif currPage == total_pages:
        currResults = total_results-20*(total_pages-1)
        RequestData(currPage)
        ShowData(currResults)
        moreBtn.grid_forget()
    MoviesData.config(state="disabled")


# Driver code
if __name__ == "__main__":

    # Create a GUI window
    root = Tk()

    # Set the background colour of GUI window
    root.config(bg="#121212", pady=20)

    # Set the configuration of GUI window
    root.geometry("1280x720")
    # root.resizable(False, False)

    # set the name of tkinter GUI window
    root.title("Movie Search")

    # Title of the window
    title = Label(root, text='Movie Search using API',
                  bg='#121212', fg='white')
    title.config(font=("Comic Sans MS", 35))
    title.pack()

    # Main Frame of the GUI
    mainFrame = Frame(root, bg='black')
    mainFrame.place(relx=.5, rely=.15, anchor="center")

    # Label for Search Term
    Label(mainFrame, text='Enter Movie Name', bg='black',
          fg='white', padx=10, font='Verdana', borderwidth=5, relief="solid").grid(row=0)

    # Text Field to enter Search Term
    searchTerm = Entry(mainFrame, font='Verdana', borderwidth=5, relief="solid")
    searchTerm.grid(row=0, column=1)

    # Creates Text Widget for showing output and calls MovieSearch() function
    def init():
        MoviesData.grid()
        moreBtn.grid_forget()
        MoviesData.config(state="normal")
        MoviesData.delete('1.0', END)
        MovieSearch()
        MoviesData.config(state="disabled")

    # Search button that calls init() Function on click
    searchBtn = Button(mainFrame, text='Search', bg='black', fg='white',
                       activebackground='cyan', font='Verdana', command=init)
    searchBtn.grid(row=0, column=2)

    # Frame for Output related widgets
    output = Frame(root, bg='#202325')
    output.place(relx=.5, rely=.6, anchor="center")

    # Text Widget to Output Movie Info
    MoviesData = Text(output, width=100, height=25, bg='black',
                      fg='white', padx=10, state="disabled", font='arial', borderwidth=2, relief="raised")
    MoviesData.grid(row=3)
    MoviesData.grid_forget()

    # More Button to show more Results
    moreBtn = Button(output, text='More', bg='black', fg='white',
                     activebackground='cyan', font='Verdana', command=loadMore)
    moreBtn.grid(row=4)
    moreBtn.grid_forget()

# Start the GUI
root.mainloop()
