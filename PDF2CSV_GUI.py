'''
Created on Jun 16, 2021

@author: Sean
'''


import tkinter as tk
import MktScan_PDF2CSV

class Convert_GUI:
    
    def __init__(self):
        '''Constructor for GUI'''
        self.input_folder = ''        #path of input PDF file
        self.output_folder = ''     #path of output CSV folder
        #text field widgets
        self.e1 = None 
        self.e2 = None
        self.e3 = None
        self.e4 = None
        #MktScan_PDF2CSV uses Stream parser of Camelot-Py Module
        self.edge_tol = ''  #how far vertically should Stream search for table
        self.row_tol = ''   #how much to group text in csv together
        self.__create_GUI()
       
        
    def __create_GUI(self):      
        root = tk.Tk(className = " Market Scan Information Systems")
        root.geometry("850x300")
 
        title = tk.Label(root, text = "Convert from PDF to Excel", font = "Helvetica 16 bold italic", fg = 'blue')
        title.place(x = 300, y = 0)

        in_file_label = tk.Label(root, text = "Input PDF Folder Path: ")
        in_file_label.place(x = 20, y = 40)

        out_file_label = tk.Label(root, text = "Output CSV Folder Path: ")
        out_file_label.place(x = 20, y = 60)

    
        self.e1 = tk.Entry(root)
        self.e2 = tk.Entry(root)
        
        self.e1.place(x = 170, y =40, width=600)   
        self.e2.place(x = 170, y= 60, width=600)



        stream_edge_label = tk.Label(root, text = "Stream edge Tolerance: ")
        stream_edge_label.place(x = 20, y = 100)
        stream_row_label = tk.Label(root, text = "Stream row tolerance: ")
        stream_row_label.place(x = 300, y = 100)

        self.e3 = tk.Entry(root)
        self.e4 = tk.Entry(root)
        self.e3.place(x = 170, y = 100, width = 100)
        self.e4.place(x = 435, y = 100, width = 100)

        sub_btn=tk.Button(root,text = 'Submit', command = self.__submit , bg = "light green", font = 'bold')
        sub_btn.place(x =100, y = 200, width = 200)
        
        root.mainloop()


    def __submit(self):
        self.input_folder = self.e1.get()
        self.output_folder = self.e2.get()
        self.edge_tol = self.e3.get()
        self.row_tol = self.e4.get()
        

        if not self.edge_tol and not self.row_tol: #no input, use default arguments
            MktScan_PDF2CSV.convert(in_path = self.input_folder, out_path = self.output_folder)
            
        elif self.edge_tol.isdigit() and self.row_tol.isdigit(): #check for valid integer
            self.edge_tol = int(self.e3.get())
            self.row_tol = int(self.e4.get())
            MktScan_PDF2CSV.convert(in_path = self.input_folder, out_path = self.output_folder,
                                 edge_tol2 = self.edge_tol, row_tol2 = self.row_tol)
            
        else:
            pop_up = tk.Tk(className = "Error - Invalid Argument")
            pop_up.geometry("400x100")
            label = tk.Label(pop_up, text = "Stream inputs must be positive integer", font = 'bold')
            label.place(x = 20, y = 40)
        
        
       







