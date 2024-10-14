import existing_ppt, new_ppt, globalvars
import datetime
import tkinter, ttkthemes, tkinter.filedialog

class Xtractor(ttkthemes.ThemedTk):
    def __init__(self):
        super().__init__()

        # configure the root window
        self.title('PowerPi')
        self.geometry('500x400')
        self.set_theme('radiance')
        # define a frame
        self.main_frame = tkinter.ttk.Frame(self)
        self.main_frame.pack()
        ## label for today's day
        self.date_label = tkinter.ttk.Label(self.main_frame, text = 'Inspection Date')
        self.date_label.grid(row = 0, column = 1)
        ## entry for today's day
        self.date_entry = tkinter.ttk.Entry(self.main_frame, width = 30, state = tkinter.NORMAL, justify = 'center')
        self.date_entry.grid(row = 1, column = 1)
        ## label for bridgeID
        self.bridgeID_label = tkinter.ttk.Label(self.main_frame, text = 'Bridge ID')
        self.bridgeID_label.grid(row = 2, column = 1)
        ## entry for bridgeID
        self.bridgeID_entry = tkinter.ttk.Entry(self.main_frame, width = 30, state = tkinter.NORMAL, justify = 'center')
        self.bridgeID_entry.grid(row = 3, column = 1)
        ## label for importing data
        self.import_label = tkinter.ttk.Label(self.main_frame, text = 'Import the field PowerPoint')
        self.import_label.grid(row = 4, column = 1)
        ## entry for importing data
        self.import_entry = tkinter.ttk.Entry(self.main_frame, width = 40, state = tkinter.NORMAL)
        self.import_entry.grid(row = 5, column = 1)
        ## button for importing data
        self.import_button = tkinter.ttk.Button(self.main_frame, text = '...', state = tkinter.NORMAL, command = self.ppt_import, width = 2)
        self.import_button.grid(row = 5, column = 2) 
        ## label for save folder
        self.save_label = tkinter.ttk.Label(self.main_frame, text = 'Select the folder where files will be saved')
        self.save_label.grid(row = 6, column = 1)
        ## entry for save folder
        self.save_entry = tkinter.ttk.Entry(self.main_frame, width = 40, state = tkinter.NORMAL)
        self.save_entry.grid(row = 7, column = 1)
        ## button for save folder
        self.save_button = tkinter.ttk.Button(self.main_frame, text = '...', state = tkinter.NORMAL, command = self.save_path,  width = 2)
        self.save_button.grid(row = 7, column = 2) 
        ## button for starting the process
        self.start_button = tkinter.ttk.Button(self.main_frame, text = 'Start', command = self.ppt_create, state = tkinter.NORMAL, width = 4)
        self.start_button.grid(row = 8, column = 1) 



    # define function that imports the Excel file
    def ppt_import(self):
        # variable that handles the Excel path
        path = tkinter.filedialog.askopenfilename(filetypes = (
            ("PowerPoint Files", "*.PPTX"),
            ("All Files", "*.*")
        ))
        # fill entry bar with the path
        self.import_entry.insert(tkinter.END, path)
        # update the file_path
        globalvars.file_path = path

        # update user_var
        count = 0
        i = 0
        j = 0
        for k in range(len(path)):
            if path[k] == "/":
                count += 1
            elif count == 2 and i == 0:
                i = k
            elif count == 3:
                j = k - 1
                break
        globalvars.user_path = path[i:j]

        # update bridgeID
        m = 0
        n = 0
        for k in range(len(path)):
            if path[k] == '0':
                m = k
                n = m + 10
                break
        globalvars.bridgeID = path[m:n]

        #update inspection_date
        globalvars.inspection_date = path[-13:-5]
        self.date_entry.insert(tkinter.END, globalvars.inspection_date)

        # fill the bridge ID bar
        self.bridgeID_entry.insert(tkinter.END, globalvars.bridgeID)


    # define function that gets the saving folder path
    def save_path(self):
        # variable that handles the Excel path
        folder_path = tkinter.filedialog.askdirectory()
        # fill entry bar with the path
        self.save_entry.insert(tkinter.END, folder_path)
        # set the path for saving files
        globalvars.save_path = folder_path


    # define function that downalods files from NBIS website
    def ppt_create(self):
        # get the date inputed by user
        globalvars.inspection_date = self.date_entry.get()
        # create file
        output_folder_pictures = globalvars.save_path + '/temp'
        image_caption_dict = existing_ppt.extract_images_and_captions_from_ppt(globalvars.file_path, output_folder_pictures)
        output_file = globalvars.save_path + f"/{globalvars.bridgeID} Routine Inspection Photos_{globalvars.inspection_date}_ARRANGED.pptx"
        globalvars.inspectors = existing_ppt.extract_inspectors_from_ppt(globalvars.file_path)
        new_ppt.create_ppt_with_two_images_per_slide(globalvars.bridgeID, globalvars.inspection_date, image_caption_dict, output_folder_pictures, output_file)


if __name__ == "__main__":
    app = Xtractor()
    app.mainloop()