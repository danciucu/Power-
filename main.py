import existing_ppt, new_ppt, new_pdf, globalvars
import tkinter, ttkthemes, tkinter.filedialog

class Xtractor(ttkthemes.ThemedTk):
    def __init__(self):
        super().__init__()

        # configure the root window
        self.title('PowerPi')
        self.geometry('500x400')
        self.set_theme('radiance')
        # set up a menu
        self.menubar = tkinter.Menu(self)
        self.menubar.add_command(label = "Settings", command = self.settings_button)
        self.config(menu = self.menubar)
        # define a frame
        self.main_frame = tkinter.ttk.Frame(self)
        self.main_frame.pack()
        ## label for today's day
        self.date_label = tkinter.ttk.Label(self.main_frame, text = 'Inspection Date (YYYYMMDD)')
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


    # define function that allows the user to change the settings
    def settings_button(self):
        # create a new window and configure the size of it
        settings_window = tkinter.Toplevel(self)
        settings_window.geometry('400x200')
        # define three variables for the radiobuttons
        batch_var = tkinter.IntVar()
        # add label for folder question
        question1_label = tkinter.ttk.Label(settings_window, text = 'Do you want to create a batch of photo documents?', justify = 'center')
        question1_label.grid(row = 0, column = 0)
        # add two radiobuttons
        yes_radiobutton = tkinter.ttk.Radiobutton(settings_window, text = 'Yes (Default)', variable = batch_var, value = 0)
        yes_radiobutton.grid(row = 1, column = 0)
        no_radiobutton = tkinter.ttk.Radiobutton(settings_window, text = 'No', variable = batch_var, value = 1)
        no_radiobutton.grid(row = 2, column = 0)
        # add label for report question
        question2_label = tkinter.ttk.Label(settings_window, text = 'Output file path', justify = 'center')
        question2_label.grid(row = 3, column = 0)
        # add an entry
        save_entry = tkinter.ttk.Entry(settings_window, width = 30, state = tkinter.NORMAL, justify = 'center')
        save_entry.grid(row = 4, column = 0)
        ## button for save folder
        save_button = tkinter.ttk.Button(settings_window, text = '...', state = tkinter.NORMAL, command = self.save_path,  width = 2)
        save_button.grid(row = 4, column = 2) 
        # define function that tracks which radiobutton the user selected
        def assign_value():
            globalvars.settings_batch = batch_var.get()
            settings_window.destroy()
        # add a button that saves the settings
        save_button = tkinter.ttk.Button(settings_window, text = "Save", command = lambda: assign_value())
        save_button.grid(row = 5, column = 0)


    # define function that imports the Excel file
    def ppt_import(self):
        # reset the entry boxes
        self.date_entry.delete(0, 'end')
        self.bridgeID_entry.delete(0, 'end')
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
        for k in range(len(path) - 3):
            if path[k].isdigit() and (path[k + 3] in ['C', 'R', 'T', 'B'] and path[k + 4].isdigit()):
                m = k
                n = m + 10
                break
        globalvars.bridgeID = path[m:n]

        # update inspection_date
        p = 0
        q = 0
        for k in range(len(path) - 4):
            if path[k:k + 4] == '2024':
                p = k
                q = k + 8
        globalvars.inspection_date = path[p:q]

        # fill the inspection date bar
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
        # create temp folder path
        output_folder_pictures = globalvars.save_path + '/temp'
        # create power point file name
        file_name_ppt = f"/{globalvars.bridgeID} Routine Inspection Photos_{globalvars.inspection_date}_ARRANGED.pptx"
        # create pdf file name
        file_name_pdf = f"/{globalvars.bridgeID} Routine Inspection Photos_{globalvars.inspection_date}_ARRANGED.pdf"
        # create power point file path
        output_file = globalvars.save_path + file_name_ppt
        # get the inspectors from existing power point
        globalvars.inspectors = existing_ppt.extract_inspectors_from_ppt(globalvars.file_path)
        # get the images from existing power point
        image_caption_dict = existing_ppt.extract_images_and_captions_from_ppt(globalvars.file_path, output_folder_pictures)
        # create new power point
        new_ppt.create_ppt_with_two_images_per_slide(globalvars.bridgeID, globalvars.inspection_date, image_caption_dict, output_folder_pictures, output_file)
        # create new pdf
        new_pdf.create_pdf(globalvars.save_path, file_name_ppt, file_name_pdf)
        # print the completion message once all files are created
        print(f'{globalvars.bridgeID} photo document completed!')


if __name__ == "__main__":
    app = Xtractor()
    app.mainloop()