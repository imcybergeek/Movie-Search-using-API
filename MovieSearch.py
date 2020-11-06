from tkinter import *
import requests
import xlsxwriter
import urllib.request
from io import BytesIO
from PIL import Image, ImageTk

# Variables containing API information which is used in making request and fetching data.
main_url = "https://api.themoviedb.org/3/search/movie?"
api = "api_key=1dcf69b9b95240032c80e5d374ca2bee"
search_query = "&query="
no_of_pages = "&page="

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
    if parsedTerm == "":
        exit()
    movie_name = parsedTerm.replace(" ", "+")
    requested_data = requests.get(url=main_url+api+search_query+movie_name)
    movie_data = requested_data.json()
    total_pages = int(movie_data["total_pages"])
    total_results = int(movie_data["total_results"])
    if (total_pages != 0):
        currPage = 1
    DataManipualation()


# Determines which set of procedures to execute based on number of results and calls ShowData() function
def DataManipualation():
    global currPage, total_pages, total_results
    if total_pages == 0:
        print("No Results Found")
        noResult= Label(outputFrame, width=67, height=2, bg="black",
                           fg="white", padx=10, font="arial", borderwidth=1, relief="raised", text="No Results Found")
        noResult.grid(row=0, column=0, pady=(20, 0))
    elif total_pages == 1:
        currResults = total_results
        ShowData(currResults)
    elif (total_pages > 1):
        ShowData(20)
        moreBtn.grid(row=0, column=3, padx=(10,0))


# Sends request to fetch more results from API and store in JSON format
def RequestData(currPage):
    global movie_name, movie_data, total_pages, total_results
    requested_data = requests.get(
        url=main_url+api+search_query+movie_name+no_of_pages+str(currPage))
    movie_data = requested_data.json()


# Extracts data by parsing JSON and feeds to Text Widget using tkinter
def ShowData(currResults):
    global movie_data, index

    for child in outputFrame.winfo_children():
        child.destroy()

    block, block2, block3 = [], [], []

    for i in range(currResults):
        index += 1
        movieName = movie_data["results"][i]["original_title"]
        overview = movie_data["results"][i]["overview"]
        image = movie_data["results"][i]["poster_path"]

        block.append(Label(outputFrame, width=50, height=2, bg="black",
                           fg="white", padx=10, font="arial", borderwidth=1, relief="raised", text=movieName))
        block[i].grid(row=2*i, column=0, pady=(20, 0))

        block2.append(Text(outputFrame, width=50, height=10, bg="black",
                           fg="white", padx=10, font="arial", wrap=WORD, borderwidth=1, relief="raised"))

        block2[i].insert(END, overview)
        block2[i].config(state="disabled")
        block2[i].grid(row=2*i+1, column=0)

        if(image != None):
            imageURL = "https://image.tmdb.org/t/p/w500"+image
            imageData = urllib.request.urlopen(imageURL)
            rawImage = imageData.read()
            imageData.close()
            image = Image.open(BytesIO(rawImage))
        else:
            image = Image.open("default.jpg")

        poster = image.resize((150, 250), Image.ANTIALIAS)
        poster = ImageTk.PhotoImage(poster)
        block3.append(Button(outputFrame, width=150, height=225, bg="black",
                            relief="raised", image=poster))
        block3[i].grid(row=2*i, column=1, rowspan=2, pady=(20, 0))
        block3[i].poster = poster

        print("\n\n Name Of Movie ", index, ": "+movieName)
        Movie_Name.append(movieName)

        if(overview != ""):
            print(" Overview Of Movie: "+overview, "\n")
            Movie_Overview.append(overview)
        else:
            print(" Overview Of Movie: NA")
            Movie_Overview.append("NA")
    SaveData(currResults)

# Saves data in Excel format with file name as movie.xlsx


def SaveData(currResults):
    global lines_XLSX
    workbook = xlsxwriter.Workbook("movie.xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.write("A1", "Movie Name")
    worksheet.write("B1", "Movie Description")
    for i in range(2, lines_XLSX+currResults):
        wb = "A"+str(i)
        wbb = "B"+str(i)
        worksheet.write(wb, Movie_Name[i-2])
        worksheet.write(wbb, Movie_Overview[i-2])
    lines_XLSX += currResults
    workbook.close()


# Loads next set of results when More button is clicked
def loadMore():
    global currPage, total_pages, total_results
    currPage += 1

    if currPage < total_pages:
        RequestData(currPage)
        ShowData(20)
    elif currPage == total_pages:
        currResults = total_results-20*(total_pages-1)
        RequestData(currPage)
        ShowData(currResults)
        moreBtn.grid_forget()

def mousewheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")

def enter(event):
    init()

# Creates Text Widget for showing output and calls MovieSearch() function
def init():
    moreBtn.grid_forget()
    MovieSearch()

# Driver code
if __name__ == "__main__":

    # Create a GUI window
    root = Tk()

    # Set the background colour of GUI window
    root.config(bg="#121212", pady=20)

    # Set the configuration of GUI window
    root.geometry("1920x1080")
    # root.resizable(False, False)

    # set the name of tkinter GUI window
    root.title("Movie Search")

    # Title of the window
    title = Label(root, text="Movie Search using API",
                  bg="#121212", fg="white")
    title.config(font=("Comic Sans MS", 35))
    title.pack()

    # Main Frame of the GUI
    mainFrame = Frame(root, bg="#121212")
    mainFrame.pack()

    # Label for Search Term
    searchLabel = Label(mainFrame, text="Enter Movie Name", bg="black",
                         fg="white", padx=10, font="Verdana", borderwidth=5, relief="solid")
    searchLabel.grid(row=0, column=0)

    # Text Field to enter Search Term
    searchTerm = Entry(mainFrame, font="Verdana",
                       borderwidth=5, relief="solid")
    searchTerm.grid(row=0, column=1)
    searchTerm.focus()

    # Search button that calls init() Function on click
    searchBtn = Button(mainFrame, text="Search", bg="black", fg="white",
                       activebackground="cyan", font="Verdana", command=init)
    searchBtn.grid(row=0, column=2)
    root.bind('<Return>', enter)

    # More Button to show more Results
    moreBtn = Button(mainFrame, text="More", bg="black", fg="white",
                     activebackground="cyan", font="Verdana", command=loadMore)
    moreBtn.grid_forget()

    output = Frame(root, bg="#121212")
    output.pack(fill=BOTH, expand=1, pady=(20,0))

    canvas = Canvas(output, bg="#121212", width=630)
    canvas.pack(side=LEFT, fill=Y, expand=1)

    canvas.configure()
    canvas.bind_all("<MouseWheel>", mousewheel)

    outputFrame = Frame(canvas, bg="#121212")
    outputFrame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0,0), window=outputFrame, anchor=NW)

root.mainloop()
