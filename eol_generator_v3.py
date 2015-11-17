#!/usr/bin/python

#
# Author: Wes MacKay
#
# To Do:
# - Modify cross-reference to be run in a thread
#
# ToDo:
# -...
#


import xlrd                      #Excel Read
import xlsxwriter                #Excel Write
import sys, os                   #Get scripts current directory
from tkinter import *            #GUI
from tkinter import ttk
import tkinter.messagebox        #Error Message Box
from tkinter import filedialog   #File Browse Dialog Box
import threading                 #Threading


# ------------------------------------------------
# Classes
# ------------------------------------------------

class createGUI:
   def __init__(self, master):
      self.master = master
      master.title("EOL Generator")       # Create title of app
      self.xcel_list = []                 # List to hold each excel file object for editing later

      ### Create Step Forms
      # First Step
      self.first_step = LabelFrame(master, text=' 1. Enter File Details: ')
      self.first_step.grid(row=0, columnspan=7, padx=5, pady=5, ipadx=5, ipady=5, sticky='WE')
      # Second Step
      self.second_step = LabelFrame(master, text=' 2. Enter Parameters: ')
      self.second_step.grid(row=1, columnspan=7, padx=5, pady=5, ipadx=5, ipady=5, sticky='WE')
      # Third Step
      self.third_step = LabelFrame(master, text=" 3. Run: ")
      self.third_step.grid(row=2, columnspan=7, padx=5, pady=5, ipadx=5, ipady=5, sticky='WE')

      ### Populate Step Forms
      # First Step
      self.i_file_path = ''
      self.i_file_lbl =  Label(self.first_step, text='Input File:')
      self.i_file_entry = Entry(self.first_step)
      self.i_file_btn = Button(self.first_step, text="Browse", command=lambda: self.browse_file("input"))
      # -------------
      self.s_file_path = ''
      self.s_file_lbl =  Label(self.first_step, text='Search File:')
      self.s_file_entry = Entry(self.first_step)
      self.s_file_btn = Button(self.first_step, text="Browse", command=lambda: self.browse_file("search"))
      # Second Step
      self.sheet_question = Label(self.second_step, text="A) Enter the sheet #'s for both files:")
      self.i_sheet = StringVar()
      self.i_sheet_lbl = Label(self.second_step, text=" --   Input Sheet #: ")
      self.i_sheet_btn = OptionMenu(self.second_step, self.i_sheet, '1', '2', '3', '4')
      self.s_sheet = StringVar()
      self.s_sheet_lbl = Label(self.second_step, text=" --   Search Sheet #: ")
      self.s_sheet_btn = OptionMenu(self.second_step, self.s_sheet, '1', '2', '3', '4')
      # -------------
      self.column_question = Label(self.second_step, text="B) Which column (A,B,C) is your part # located?")
      self.i_column = StringVar()
      self.i_column_lbl = Label(self.second_step, text=" --   Input Column: ")
      self.i_column_btn = OptionMenu(self.second_step, self.i_column, 'A', 'B', 'C', 'D')
      self.s_column = StringVar()
      self.s_column_lbl = Label(self.second_step, text=" --   Search Column: ")
      self.s_column_btn = OptionMenu(self.second_step, self.s_column, 'A', 'B', 'C', 'D')
      # -------------
      self.columns_to_keep_lbl = Label(self.second_step, text="C) Enter the columns (A,B,C) to keep in the output:")
      self.i_columns_to_keep_lbl = Label(self.second_step, text=" --   Input Columns: ")
      self.i_columns_to_keep = Entry(self.second_step)
      self.s_columns_to_keep_lbl = Label(self.second_step, text=" --   Search Columns: ")
      self.s_columns_to_keep = Entry(self.second_step)
      # Third Step
      self.load_params_btn = Button(self.third_step, text='Pre-Load Config', command=self.preload)
      self.submit_btn = Button(self.third_step, text='Start Searching!', command=self.get_input)
      self.progressbar = ttk.Progressbar(self.third_step, orient=HORIZONTAL, length=270, mode='determinate')

      ### Assign things to GUI GRID
      # First Step
      self.i_file_lbl.grid(row=0, column=0, padx=5, pady=2, sticky='E')
      self.i_file_entry.grid(row=0, column=1, columnspan=7, pady=3, sticky="WE")
      self.i_file_btn.grid(row=0, column=8, padx=5, pady=2, sticky='E')
      # -------------
      self.s_file_lbl.grid(row=1, column=0, padx=5, pady=2, sticky='E')
      self.s_file_entry.grid(row=1, column=1, columnspan=7, pady=3, sticky="WE")
      self.s_file_btn.grid(row=1, column=8, padx=5, pady=2, sticky='E')
      # Second Step
      self.sheet_question.grid(row=0, column=0, padx=5, pady=2, sticky='W')
      self.i_sheet_lbl.grid(row=1, column=0, padx=5, pady=2, sticky='W')
      self.i_sheet_btn.grid(row=1, column=0, padx=5, pady=2, sticky='E')
      self.s_sheet_lbl.grid(row=2, column=0, padx=5, pady=2, sticky='W')
      self.s_sheet_btn.grid(row=2, column=0, padx=5, pady=2, sticky='E')
      # -------------
      self.column_question.grid(row=3, column=0, padx=5, pady=2, sticky='W')
      self.i_column_lbl.grid(row=4, column=0, padx=5, pady=2, sticky='W')
      self.i_column_btn.grid(row=4, column=0, padx=5, pady=2, sticky='E')
      self.s_column_lbl.grid(row=5, column=0, padx=5, pady=2, sticky='W')
      self.s_column_btn.grid(row=5, column=0, padx=5, pady=2, sticky='E')
      # -------------
      self.columns_to_keep_lbl.grid(row=6, column=0, padx=5, pady=2, sticky='W')
      self.i_columns_to_keep_lbl.grid(row=7, column=0, padx=5, pady=2, sticky='W')
      self.i_columns_to_keep.grid(row=7, column=0, padx=5, pady=2, sticky='E')
      self.s_columns_to_keep_lbl.grid(row=8, column=0, padx=5, pady=2, sticky='W')
      self.s_columns_to_keep.grid(row=8, column=0, padx=5, pady=2, sticky='E')
      # Third Step
      self.load_params_btn.grid(row=0, column=0, padx=5, pady=2, sticky='W')
      self.submit_btn.grid(row=0, column=2, padx=5, pady=2, sticky='E')
      self.progressbar.grid(row=1, columnspan=3, padx=5, pady=2, sticky='WE')


   # Get input from forms
   def get_input(self):
      all_parameters = [self.i_file_path, self.i_sheet.get(), self.i_column.get(), self.i_columns_to_keep.get(), self.s_file_path, self.s_sheet.get(), self.s_column.get(), self.s_columns_to_keep.get()]
      # Check if all variables were entered
      if not all(all_parameters):
         tkinter.messagebox.showinfo('Not enough info', 'Please answer every question. Try again.')
      else:
         # Validate if the files are of file types
         if not self.check_files([self.i_file_path, self.s_file_path]): return
         # Convert columns from (A,B,C) to (0,1,2)
         for i in [2, 3, 6, 7]: all_parameters[i] = convert_columns(all_parameters[i])
         # Disable submit button
         self.submit_btn.config(state="disabled")
         self.load_params_btn.config(state="disabled")
         # Flag to indicate status of thread completion
         # 0 - nothing finished / 1 - file threads finished
         self.flag = 0
         # Create separate threads for opening each xcel file
         self.i_thread = ThreadedClient("input", all_parameters, self.xcel_list)
         self.s_thread = ThreadedClient("search", all_parameters, self.xcel_list)
         self.o_thread = ThreadedClient("output", all_parameters, self.xcel_list)
         # Start Threads
         self.i_thread.start()
         self.s_thread.start()
         self.o_thread.start()
         # Wait for threads to finish
         self.wait_for_threads()
         self.progressbar.start(50)

   # Wait until threads are still running, and then execute cross_reference function
   def wait_for_threads(self):
      # If background threads are still running, keep waiting until they finish
      if threading.active_count() > 1:
         self.master.after(100, self.wait_for_threads)
      # If file threads have just completed
      elif self.flag == 0:
         # Check files for any errors
         error = self.check_for_errors()
         # If there was no errors, cross-reference values
         if error is False:
            # Start cross_reference function in a separate thread
            search_thread = threading.Thread(target = cross_reference, args = [self.xcel_list])
            search_thread.start()
            # Set flag to represent file threads have completed
            self.flag = 1
            # Restart wait_for_threads function to wait for search_thread to finish
            self.master.after(100, self.wait_for_threads)
         # If there was an error, stop and reset GUI and values
         else: self.cleanup()
      # If the search_thread is finished, stop and reset GUI and values
      elif self.flag == 1: self.cleanup()

   # Cleanup/clear GUI values for another run
   def cleanup(self):
      # Stop Progressbar
      self.progressbar.stop()
      # Save changes to output_file worksheet and close file
      output = self.xcel_list[-1]
      close(output)
      # Clear xcel list for next run
      self.xcel_list = []
      # Restore submit button to active state
      self.submit_btn.config(state="active")
      self.load_params_btn.config(state="active")

   # Checks xcel_list for any file filled with 'None' type, if so return error = True
   def check_for_errors(self):
      # Check if all threads finished successfully, if there was an error tell user
      error = False
      for file in self.xcel_list:
         # If there was error when opening a file, tell user
         if file.file == None:
            tkinter.messagebox.showinfo('Error', "File \'%s\' could not be opened" % file.file_name)
            error = True
      return error

   # Validate if files exist
   def check_files(self, files):
      for file in files:
         try:
            with open(file) as f:
               pass
         except IOError as e:
            tkinter.messagebox.showinfo('Error', "File \'%s\' could not be opened" % file)
            return False
      return True

   # Open Browse button to locate input/search files
   def browse_file(self, type):
      # Open Browse Dialog
      file_path = filedialog.askopenfilename(initialdir=os.getcwd())    # Full path
      file = file_path.split('/')[-1]                                   # Only file name
      if type == "input":
         self.update(self.i_file_entry, file, "entry")
         self.i_file_path = file_path
      elif type == "search":
         self.update(self.s_file_entry, file, "entry")
         self.s_file_path = file_path

   # Preload Parameters for testing
   def preload(self):
      # If user has not entered their own file, load our presets
      if self.i_file_entry.get() == '':
         self.update(self.i_file_entry, "input_file.xlsx", "entry")
         self.i_file_path = os.getcwd() + '\input_file.xlsx'
      if self.s_file_entry.get() == '':
         self.update(self.s_file_entry, "search_file.xlsx", "entry")
         self.s_file_path = os.getcwd() + '\search_file.xlsx'
      self.update(self.i_sheet, "1", "optionmenu")
      self.update(self.s_sheet, "1", "optionmenu")
      self.update(self.i_column, "C", "optionmenu")
      self.update(self.s_column, "A", "optionmenu")
      self.update(self.i_columns_to_keep, "A,B,C,D", "entry")
      self.update(self.s_columns_to_keep, "O,P,Q", "entry")

   # Update values on GUI
   def update(self, entry, value, type):
      if type == "entry":
         entry.delete(0, END)
         entry.insert(0, value)
      elif type == "optionmenu":
         entry.set(value)


# Load excel file with it's attributes into a class object
class ThreadedClient(threading.Thread):
   def __init__(self, file, all_parameters, xcel_list):
      threading.Thread.__init__(self)
      self.file = file
      self.all_parameters = all_parameters
      self.xcel_list = xcel_list
   # Create individual xcel objects to modify data later...and add them to global xcel_list
   def run(self):
      this_file = xcel_file(*sort_parameters(self.file, self.all_parameters))
      # Add this_file to xcel_list, to process later
      if self.file == "input":
         self.xcel_list.insert(0, this_file)
      elif self.file == "search":
         self.xcel_list.insert(1, this_file)
      elif self.file == "output":
         self.xcel_list.append(this_file)


# Class to hold each excel file + their values
class xcel_file(object):
   # Variables unique to specific call
   def __init__(self, file, sheet_num, search_column = '', columns_to_keep = ''):
      # Specific operations needed for the output file
      if file == 'output_file.xlsx':
         self.file_name = file
         self.file = xlsxwriter.Workbook('output_file.xlsx')
         self.sheet = self.file.add_worksheet("Output Values")
         self.parameters = [file, sheet_num]
      # Operations for the input & search files
      else:
         self.file_name = file
         self.open(file, sheet_num)
         self.search_column = int(search_column)
         self.columns_to_keep = columns_to_keep
         self.parameters = [file, sheet_num, search_column, columns_to_keep]
   # Try to open given file, or return file = None if it fails
   def open(self, this_file, this_sheet_num):
         try:
            self.file = xlrd.open_workbook(this_file)
            self.sheet = self.file.sheet_by_index(int(this_sheet_num)-1)
         except Exception as e:
            print("Error: File \'%s\' could not be opened" % this_file)
            # Set file = None so we can determine if this thread has failed to complete
            self.file = None


# Close out given file
def close(this_file):
   try:
      this_file.file.close()
   except:
      print("\nError: Could not close \"%s\"", this_file.file_name)


# Split up full_parameters and then return only the relevant args according to 'name'
def sort_parameters(name, full_parameters):
   # Split up parameters to reflect which params belong to the 'name' file
   # [i_file, i_sheet, i_column, i_columns_to_keep, s_file, s_sheet, s_column, s_columns_to_keep, o_file, o_sheet]
   if name == 'input':
      parameters = cut(full_parameters, [0,1,2,3])
   elif name == 'search':
      parameters = cut(full_parameters, [4,5,6,7])
   elif name == 'output':
      parameters = ["output_file.xlsx", "Output Values"]
      #parameters = cut(full_parameters, [9,10])
   # Return the 5 parameters for the xcel class object
   return parameters


# Convert (A,B,C) values to (1,2,3)
def convert_columns(initial_columns):
   # Convert letters to numbers
   columns = [ord(char) - 96 for char in initial_columns.lower()]
   # Only keep 1-27 values (alphabet)...remove ',' & ' ' values
   # (x-1) is needed for 0 index in dealing with xcel spreadsheets
   columns = [(x-1) for x in columns if x in range(1, 27)]
   # If list only contains 1 element, convert list to value
   if len(initial_columns) == 1:
      columns = columns[0]
   return columns


# Search given worksheet for string and return full row that is found
def grep(string, worksheet, file):
   # Loop through search_file line by line
   for row in range(worksheet.nrows):
      # Grep command - We know part_num in search_file is in the A column
      if re.match("^" + string + "$", str(worksheet.cell_value(row, file.search_column))):
         return worksheet.row(row)
   # If nothing is found, return empty string
   return ""


# Remove everything except specified columns in line
def cut(line, columns):
   list = []
   for col in columns:
      cell = line[col]
      list.append(cell)
   return list


# Prints line to output_file
def printout(line, row, file):
   # Analyze line for value types and print to output_file
   for i, cell in enumerate(line):
      # Check if cell is a string (not object of string)
      if isinstance(cell, str):
         file.sheet.write(row, i, cell)
      # Convert date value to readable format ('3' = date format)
      elif cell.ctype == 3:
         file.sheet.write(row, i, cell.value, date_format)
      else:
         file.sheet.write(row, i, str(cell.value))


# Cross-Reference the Excel Spreadsheets
def cross_reference(xcel_list):
   # Set xcel files to local variables
   input = xcel_list[0]
   search = xcel_list[1]
   output = xcel_list[-1]

   ## Cross-Reference part_nums and store the unique values
   similar_values = set(input.sheet.col_values(input.search_column)) & set(search.sheet.col_values(search.search_column))
   if not len(similar_values) > 0:
      tkinter.messagebox.showinfo('No values found', 'There are no values that match.')
      close(output)
      return

   ## ------------------------------------
   ## Apply default formats to output file
   ## ------------------------------------
   # Create date format for date values (for excel date format)
   global date_format
   date_format = output.file.add_format({'num_format': 'mm/dd/yyyy', 'align': 'center'})
   bold_format = output.file.add_format({'bold': True, 'align': 'center'})
   center_format = output.file.add_format({'align': 'center'})
   # Apply a bold format to the first row
   output.sheet.set_row(0, None, bold_format)
   # Apply columns A:G with a center format
   output.sheet.set_column('A:G', None, center_format)
   # Apply a certain width to our columns
   output.sheet.set_column('A:C', 20)
   output.sheet.set_column('D:D', 35)
   output.sheet.set_column('E:G', 12)

   ## ----------------------------
   ## Read input file line by line
   ## ----------------------------
   for row_id in range(0, input.sheet.nrows):
      # This line in input_file
      line = input.sheet.row(row_id)
      # Get part_num from this line (column is set above)
      part_num = line[input.search_column].value

      # If value isn't found in search file, print existing values to output_file
      if part_num not in similar_values or row_id == 0:
         # If first line, print existing header text + wanted header text @ top of file
         if row_id == 0:
            printout(line, row_id, output)
         else:
            line = cut(line, input.columns_to_keep)
            printout(line, row_id, output)
         continue

      # Only keep wanted columns from full line
      line = cut(line, input.columns_to_keep)
      # Find our Part #'s line inside search_file
      s_line = grep(part_num, search.sheet, search)
      # Cut out only wanted_values from full line
      s_values = cut(s_line, search.columns_to_keep)
      # Merge Description values
      #if line[-1].value == s_values[0].value or s_values[0].value != '':
         # Remove description column from input line
      #   line.pop()
      # Write values to output_file
      line.extend(s_values)
      printout(line, row_id, output)

   # Display Success Window notifying user
   message = 'Found ' + str(len(similar_values)) + ' matches. Exporting results to the output file.'
   tkinter.messagebox.showinfo('Success', message)


# ------------------------------------------------
# Main Function
# ------------------------------------------------

if __name__ == '__main__':
   root = Tk()
   # Create GUI
   my_app = createGUI(root)
   # Loop GUI until exit button is pressed
   root.mainloop()