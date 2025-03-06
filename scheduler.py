import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
from datetime import datetime, timedelta
import sqlite3
import re
import os

class SchedulerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Work Schedule Manager")
        self.root.geometry("1200x800")
        
        # start the database connection
        self.db_file = 'data/schedule.db'
        self.ensure_database_exists()
        
        # make the main container
        self.main_container = ttk.Frame(root)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # make a notebook for tabs
        self.notebook = ttk.Notebook(self.main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # make the tabs
        self.setup_workplace_tab()
        self.setup_import_tab()
        self.setup_schedule_tab()
        self.setup_shift_management_tab()
        
        # update the workplace dropdown lists
        self.update_workplace_list()

    def ensure_database_exists(self):
        if not os.path.exists('data'):
            os.makedirs('data')
            
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        
        # make workplaces table
        c.execute('''CREATE TABLE IF NOT EXISTS workplaces
                    (id INTEGER PRIMARY KEY,
                     name TEXT NOT NULL,
                     hours_open TEXT NOT NULL,
                     hours_close TEXT NOT NULL)''')
        
        # make the workers table with work study field
        c.execute('''CREATE TABLE IF NOT EXISTS workers
                    (id INTEGER PRIMARY KEY,
                     workplace_id INTEGER,
                     first_name TEXT NOT NULL,
                     last_name TEXT NOT NULL,
                     email TEXT NOT NULL,
                     work_study BOOLEAN NOT NULL,
                     preferred_shifts INTEGER DEFAULT 0,
                     FOREIGN KEY (workplace_id) REFERENCES workplaces(id))''')
        
        # make availability table
        c.execute('''CREATE TABLE IF NOT EXISTS availability
                    (id INTEGER PRIMARY KEY,
                     worker_id INTEGER,
                     day TEXT NOT NULL,
                     start_time TEXT NOT NULL,
                     end_time TEXT NOT NULL,
                     FOREIGN KEY (worker_id) REFERENCES workers(id))''')
        
        # make schedules table
        c.execute('''CREATE TABLE IF NOT EXISTS schedules
                    (id INTEGER PRIMARY KEY,
                     workplace_id INTEGER,
                     worker_id INTEGER,
                     date TEXT NOT NULL,
                     start_time TEXT NOT NULL,
                     end_time TEXT NOT NULL,
                     FOREIGN KEY (workplace_id) REFERENCES workplaces(id),
                     FOREIGN KEY (worker_id) REFERENCES workers(id))''')
        
        # make shifts table
        c.execute('''CREATE TABLE IF NOT EXISTS shifts
                    (id INTEGER PRIMARY KEY,
                     workplace_id INTEGER,
                     day TEXT NOT NULL,
                     start_time TEXT NOT NULL,
                     end_time TEXT NOT NULL,
                     positions INTEGER DEFAULT 1,
                     FOREIGN KEY (workplace_id) REFERENCES workplaces(id))''')
                     
        conn.commit()
        conn.close()
        
    def setup_workplace_tab(self):
        workplace_frame = ttk.Frame(self.notebook)
        self.notebook.add(workplace_frame, text="Workplace Management")
        
        # workplace controls
        ttk.Label(workplace_frame, text="Workplace Name:").grid(row=0, column=0, pady=5, sticky='w')
        self.workplace_name = ttk.Entry(workplace_frame, width=30)
        self.workplace_name.grid(row=0, column=1, pady=5, sticky='w')
        
        ttk.Label(workplace_frame, text="Opening Time (HH:MM AM/PM):").grid(row=1, column=0, pady=5, sticky='w')
        self.opening_time = ttk.Entry(workplace_frame, width=30)
        self.opening_time.grid(row=1, column=1, pady=5, sticky='w')
        self.opening_time.insert(0, "12:00 PM")
        
        ttk.Label(workplace_frame, text="Closing Time (HH:MM AM/PM):").grid(row=2, column=0, pady=5, sticky='w')
        self.closing_time = ttk.Entry(workplace_frame, width=30)
        self.closing_time.grid(row=2, column=1, pady=5, sticky='w')
        self.closing_time.insert(0, "12:00 AM")
        
        ttk.Button(workplace_frame, text="Add Workplace", 
                  command=self.add_workplace).grid(row=3, column=0, columnspan=1, pady=10, sticky='w')
        
        ttk.Button(workplace_frame, text="Delete Selected Workplace", 
                  command=self.delete_workplace).grid(row=3, column=1, columnspan=1, pady=10, sticky='w')
        
        # workplace list
        self.workplace_list = ttk.Treeview(workplace_frame, columns=("id", "Name", "Hours"), show="headings")
        self.workplace_list.grid(row=4, column=0, columnspan=2, pady=10, sticky='nsew')
        self.workplace_list.heading("id", text="ID")
        self.workplace_list.heading("Name", text="Name")
        self.workplace_list.heading("Hours", text="Operating Hours")
        self.workplace_list.column("id", width=50)
        self.workplace_list.column("Name", width=200)
        self.workplace_list.column("Hours", width=200)
        
        # make the treeview expand with the window
        workplace_frame.columnconfigure(0, weight=1)
        workplace_frame.columnconfigure(1, weight=1)
        workplace_frame.rowconfigure(4, weight=1)
        
    def setup_import_tab(self):
        import_frame = ttk.Frame(self.notebook)
        self.notebook.add(import_frame, text="Import Workers")
        
        # control frame
        control_frame = ttk.Frame(import_frame)
        control_frame.pack(fill=tk.X, pady=10)
        
        # workplace selection
        ttk.Label(control_frame, text="Select Workplace:").pack(side=tk.LEFT, padx=5)
        self.workplace_var = tk.StringVar()
        self.workplace_dropdown = ttk.Combobox(control_frame, textvariable=self.workplace_var, width=30)
        self.workplace_dropdown.pack(side=tk.LEFT, padx=5)
        
        # excel import
        ttk.Button(control_frame, text="Import Excel File", 
                  command=self.import_excel).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="Clear All Workers", 
                  command=self.clear_workers).pack(side=tk.LEFT, padx=5)
        
        # worker list and availability frame
        workers_frame = ttk.Frame(import_frame)
        workers_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # split the frame horizontally
        workers_frame.columnconfigure(0, weight=2)
        workers_frame.columnconfigure(1, weight=3)
        workers_frame.rowconfigure(0, weight=1)
        
        # work list panel
        worker_list_frame = ttk.LabelFrame(workers_frame, text="Workers")
        worker_list_frame.grid(row=0, column=0, sticky='nsew', padx=5)
        
        self.worker_list = ttk.Treeview(worker_list_frame, 
                                      columns=("id", "Name", "Email", "Work Study"),
                                      show="headings")
        self.worker_list.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
        self.worker_list.heading("id", text="ID")
        self.worker_list.heading("Name", text="Name")
        self.worker_list.heading("Email", text="Email")
        self.worker_list.heading("Work Study", text="Work Study")
        self.worker_list.column("id", width=50)
        
        self.worker_list.bind('<<TreeviewSelect>>', self.show_worker_availability)
        
        # availability panel
        availability_frame = ttk.LabelFrame(workers_frame, text="Availability")
        availability_frame.grid(row=0, column=1, sticky='nsew', padx=5)
        
        self.availability_list = ttk.Treeview(availability_frame, 
                                           columns=("Day", "Start Time", "End Time"),
                                           show="headings")
        self.availability_list.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
        self.availability_list.heading("Day", text="Day")
        self.availability_list.heading("Start Time", text="Start Time")
        self.availability_list.heading("End Time", text="End Time")
        
    def setup_schedule_tab(self):
        schedule_frame = ttk.Frame(self.notebook)
        self.notebook.add(schedule_frame, text="Generate Schedule")
        
        # controls frame
        controls = ttk.Frame(schedule_frame)
        controls.pack(fill=tk.X, pady=10)
        
        ttk.Label(controls, text="Select Workplace:").pack(side=tk.LEFT, padx=5)
        self.schedule_workplace_var = tk.StringVar()
        self.schedule_workplace_dropdown = ttk.Combobox(controls, 
                                                      textvariable=self.schedule_workplace_var,
                                                      width=30)
        self.schedule_workplace_dropdown.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(controls, text="Generate Schedule", 
                  command=self.generate_schedule).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(controls, text="Export Schedule", 
                  command=self.export_schedule).pack(side=tk.LEFT, padx=5)
        
        # date selection
        ttk.Label(controls, text="Start Date (YYYY-MM-DD):").pack(side=tk.LEFT, padx=5)
        self.start_date_entry = ttk.Entry(controls, width=15)
        self.start_date_entry.pack(side=tk.LEFT, padx=5)
        
        # get current date and set as default
        today = datetime.now().strftime("%Y-%m-%d")
        self.start_date_entry.insert(0, today)
        
        # schedule display
        schedule_display_frame = ttk.Frame(schedule_frame)
        schedule_display_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # create scrollbars
        x_scrollbar = ttk.Scrollbar(schedule_display_frame, orient=tk.HORIZONTAL)
        y_scrollbar = ttk.Scrollbar(schedule_display_frame, orient=tk.VERTICAL)
        
        # create treeview with scrollbars
        self.schedule_display = ttk.Treeview(schedule_display_frame, 
                                         columns=("Time", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"),
                                         show="headings",
                                         xscrollcommand=x_scrollbar.set,
                                         yscrollcommand=y_scrollbar.set)
        
        # configure scrollbars
        x_scrollbar.config(command=self.schedule_display.xview)
        y_scrollbar.config(command=self.schedule_display.yview)
        
        # pack scrollbars and treeview
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.schedule_display.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        for col in ("Time", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"):
            self.schedule_display.heading(col, text=col)
            self.schedule_display.column(col, width=150, minwidth=100)
    
    def setup_shift_management_tab(self):
        shift_frame = ttk.Frame(self.notebook)
        self.notebook.add(shift_frame, text="Shift Management")
        
        # top controls
        controls = ttk.Frame(shift_frame)
        controls.pack(fill=tk.X, pady=10)
        
        ttk.Label(controls, text="Select Workplace:").pack(side=tk.LEFT, padx=5)
        self.shift_workplace_var = tk.StringVar()
        self.shift_workplace_dropdown = ttk.Combobox(controls, 
                                                  textvariable=self.shift_workplace_var,
                                                  width=30)
        self.shift_workplace_dropdown.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(controls, text="Load Shifts", 
                  command=self.load_shifts).pack(side=tk.LEFT, padx=5)
        
        # shift adding controls
        add_frame = ttk.LabelFrame(shift_frame, text="Add New Shift")
        add_frame.pack(fill=tk.X, pady=10, padx=10)
        
        # day selection
        ttk.Label(add_frame, text="Day:").grid(row=0, column=0, pady=5, padx=5, sticky='w')
        self.shift_day_var = tk.StringVar()
        day_options = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        self.shift_day_dropdown = ttk.Combobox(add_frame, 
                                            textvariable=self.shift_day_var,
                                            values=day_options,
                                            width=15)
        self.shift_day_dropdown.grid(row=0, column=1, pady=5, padx=5, sticky='w')
        
        # start time
        ttk.Label(add_frame, text="Start Time (HH:MM AM/PM):").grid(row=0, column=2, pady=5, padx=5, sticky='w')
        self.shift_start_time = ttk.Entry(add_frame, width=15)
        self.shift_start_time.grid(row=0, column=3, pady=5, padx=5, sticky='w')
        self.shift_start_time.insert(0, "12:00 PM")
        
        # end time
        ttk.Label(add_frame, text="End Time (HH:MM AM/PM):").grid(row=0, column=4, pady=5, padx=5, sticky='w')
        self.shift_end_time = ttk.Entry(add_frame, width=15)
        self.shift_end_time.grid(row=0, column=5, pady=5, padx=5, sticky='w')
        self.shift_end_time.insert(0, "5:00 PM")
        
        # positions
        ttk.Label(add_frame, text="Positions:").grid(row=1, column=0, pady=5, padx=5, sticky='w')
        self.shift_positions = ttk.Spinbox(add_frame, from_=1, to=10, width=5)
        self.shift_positions.grid(row=1, column=1, pady=5, padx=5, sticky='w')
        
        # add button
        ttk.Button(add_frame, text="Add Shift", 
                  command=self.add_shift).grid(row=1, column=2, columnspan=2, pady=5, padx=5, sticky='w')
        
        # shift list
        shift_list_frame = ttk.Frame(shift_frame)
        shift_list_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=10)
        
        self.shift_list = ttk.Treeview(shift_list_frame, 
                                     columns=("id", "Day", "Start Time", "End Time", "Positions"),
                                     show="headings")
        self.shift_list.pack(fill=tk.BOTH, expand=True)
        
        self.shift_list.heading("id", text="ID")
        self.shift_list.heading("Day", text="Day")
        self.shift_list.heading("Start Time", text="Start Time")
        self.shift_list.heading("End Time", text="End Time")
        self.shift_list.heading("Positions", text="Positions")
        
        self.shift_list.column("id", width=50)
        
        # delete button
        ttk.Button(shift_list_frame, text="Delete Selected Shift", 
                  command=self.delete_shift).pack(pady=5)
    
    def add_workplace(self):
        name = self.workplace_name.get()
        open_time = self.opening_time.get()
        close_time = self.closing_time.get()
        
        if not all([name, open_time, close_time]):
            messagebox.showerror("Error", "Please fill all fields")
            return
            
        try:
            # validate time format
            try:
                datetime.strptime(open_time, "%I:%M %p")
                datetime.strptime(close_time, "%I:%M %p")
            except ValueError:
                messagebox.showerror("Error", "Time must be in format 'HH:MM AM/PM' (e.g., '9:00 AM')")
                return
                
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            c.execute('''INSERT INTO workplaces (name, hours_open, hours_close)
                        VALUES (?, ?, ?)''', (name, open_time, close_time))
            conn.commit()
            conn.close()
            
            self.update_workplace_list()
            self.workplace_name.delete(0, tk.END)
            
            messagebox.showinfo("Success", f"Workplace '{name}' added successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add workplace: {str(e)}")
    
    def delete_workplace(self):
        selected = self.workplace_list.selection()
        if not selected:
            messagebox.showerror("Error", "Please select a workplace to delete")
            return
            
        workplace_id = self.workplace_list.item(selected[0], "values")[0]
        
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this workplace? This will remove all workers and schedules associated with it."):
            try:
                conn = sqlite3.connect(self.db_file)
                c = conn.cursor()
                
                # delete all related data first
                c.execute("DELETE FROM schedules WHERE workplace_id = ?", (workplace_id,))
                c.execute("DELETE FROM shifts WHERE workplace_id = ?", (workplace_id,))
                
                # get worker IDs to delete their availability
                c.execute("SELECT id FROM workers WHERE workplace_id = ?", (workplace_id,))
                worker_ids = c.fetchall()
                
                for worker_id in worker_ids:
                    c.execute("DELETE FROM availability WHERE worker_id = ?", (worker_id[0],))
                
                # delete workers
                c.execute("DELETE FROM workers WHERE workplace_id = ?", (workplace_id,))
                
                # finally delete the workplace
                c.execute("DELETE FROM workplaces WHERE id = ?", (workplace_id,))
                
                conn.commit()
                conn.close()
                
                self.update_workplace_list()
                messagebox.showinfo("Success", "Workplace deleted successfully!")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete workplace: {str(e)}")
    
    def update_workplace_list(self):
        # clear existing items
        for item in self.workplace_list.get_children():
            self.workplace_list.delete(item)
            
        # fetch and display workplaces
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('SELECT id, name, hours_open, hours_close FROM workplaces')
        workplaces = c.fetchall()
        conn.close()
        
        for workplace in workplaces:
            hours_str = f"{workplace[2]} - {workplace[3]}"
            self.workplace_list.insert('', 'end', values=(workplace[0], workplace[1], hours_str))
            
        # update the workplace dropdowns
        workplace_names = [w[1] for w in workplaces]
        workplace_ids = [w[0] for w in workplaces]
        
        # create a mapping of workplace names to IDs
        self.workplace_id_map = dict(zip(workplace_names, workplace_ids))
        
        self.workplace_dropdown['values'] = workplace_names
        self.schedule_workplace_dropdown['values'] = workplace_names
        self.shift_workplace_dropdown['values'] = workplace_names
    
    def import_excel(self):
        if not self.workplace_var.get():
            messagebox.showerror("Error", "Please select a workplace first")
            return
            
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if not filename:
            return
            
        try:
            # read from Excel file
            df = pd.read_excel(filename)
            
            # get the workplace_id
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            c.execute('SELECT id FROM workplaces WHERE name = ?', 
                     (self.workplace_var.get(),))
            workplace_id = c.fetchone()[0]
            
            # import the workers and their availability
            imported_count = 0
            for _, row in df.iterrows():
                # check if we have the expected columns
                required_columns = ['First Name', 'Last Name', 'Email', 'Work Study']
                if not all(col in row.index for col in required_columns):
                    messagebox.showerror("Error", "Excel file missing required columns. Expected: First Name, Last Name, Email, Work Study")
                    conn.close()
                    return
                
                # converting work study value
                work_study_val = False
                if pd.notna(row['Work Study']):
                    work_study_val = row['Work Study'].upper() == 'Y'
                
                # adding workers
                c.execute('''INSERT INTO workers 
                            (workplace_id, first_name, last_name, email, work_study)
                            VALUES (?, ?, ?, ?, ?)''',
                         (workplace_id, row['First Name'], row['Last Name'], 
                          row['Email'], work_study_val))
                worker_id = c.lastrowid
                imported_count += 1
                
                # adding availability
                for day in ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 
                           'Thursday', 'Friday', 'Saturday']:
                    if day in row and pd.notna(row[day]) and row[day].lower() != 'na':
                        try:
                            start, end = self.parse_time_range(row[day])
                            c.execute('''INSERT INTO availability 
                                        (worker_id, day, start_time, end_time)
                                        VALUES (?, ?, ?, ?)''',
                                     (worker_id, day, start, end))
                        except Exception as e:
                            messagebox.showwarning("Warning", f"Could not parse time for {row['First Name']} {row['Last Name']} on {day}: {row[day]}")
            
            conn.commit()
            conn.close()
            
            self.update_worker_list()
            messagebox.showinfo("Success", f"{imported_count} workers imported successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import Excel file: {str(e)}")
    
    def parse_time_range(self, time_str):
        # convert "2 pm - 12 am" format to proper 12 hour times
        # strip any extra whitespace
        time_str = time_str.strip()
        
        try:
            # split by hyphen or dash
            if '-' in time_str:
                parts = time_str.split('-')
            else:
                # if no hyphen, try alternative delimiter
                parts = re.split(r'\s+to\s+', time_str)
            
            start_part = parts[0].strip().lower()
            end_part = parts[1].strip().lower()
            
            # check if am/pm is missing and add it
            if 'am' not in start_part and 'pm' not in start_part:
                if int(re.search(r'(\d+)', start_part).group(1)) < 12:
                    start_part += ' am'
                else:
                    start_part += ' pm'
                    
            if 'am' not in end_part and 'pm' not in end_part:
                if int(re.search(r'(\d+)', end_part).group(1)) < 12:
                    end_part += ' am'
                else:
                    end_part += ' pm'
            
            # try different time formats
            formats = [
                '%I %p',      # e.g., "2 pm"
                '%I:%M %p',   # e.g., "2:00 pm"
                '%I%p',       # e.g., "2pm"
                '%I.%M %p'    # e.g., "2.00 pm"
            ]
            
            start_time = None
            end_time = None
            
            # try each format for start time
            for fmt in formats:
                try:
                    start_time = datetime.strptime(start_part, fmt)
                    break
                except ValueError:
                    continue
                    
            # try each format for end time
            for fmt in formats:
                try:
                    end_time = datetime.strptime(end_part, fmt)
                    break
                except ValueError:
                    continue
            
            if start_time is None or end_time is None:
                raise ValueError(f"Could not parse time range: {time_str}")
                
            # format to standard format
            start = start_time.strftime('%I:%M %p')
            end = end_time.strftime('%I:%M %p')
            
            return start, end
            
        except Exception as e:
            raise ValueError(f"Error parsing time range '{time_str}': {str(e)}")
    
    def update_worker_list(self):
        # clear all existing items
        for item in self.worker_list.get_children():
            self.worker_list.delete(item)
            
        if not self.workplace_var.get():
            return
            
        # fetch and display workers
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('''SELECT w.id, w.first_name, w.last_name, w.email, w.work_study 
                    FROM workers w
                    JOIN workplaces p ON w.workplace_id = p.id 
                    WHERE p.name = ?''', (self.workplace_var.get(),))
        workers = c.fetchall()
        conn.close()
        
        for worker in workers:
            self.worker_list.insert('', 'end', 
                                  values=(worker[0],
                                        f"{worker[1]} {worker[2]}", 
                                        worker[3], 
                                        'Yes' if worker[4] else 'No'))
    
    def show_worker_availability(self, event):
        # clear availability list
        for item in self.availability_list.get_children():
            self.availability_list.delete(item)
            
        selected = self.worker_list.selection()
        if not selected:
            return
            
        worker_id = self.worker_list.item(selected[0], "values")[0]
        
        # fetch and display availability
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('''SELECT day, start_time, end_time
                    FROM availability
                    WHERE worker_id = ?
                    ORDER BY CASE 
                        WHEN day = 'Sunday' THEN 1
                        WHEN day = 'Monday' THEN 2
                        WHEN day = 'Tuesday' THEN 3
                        WHEN day = 'Wednesday' THEN 4
                        WHEN day = 'Thursday' THEN 5
                        WHEN day = 'Friday' THEN 6
                        WHEN day = 'Saturday' THEN 7
                    END''', (worker_id,))
        availability = c.fetchall()
        conn.close()
        
        for avail in availability:
            self.availability_list.insert('', 'end', values=avail)
    
    def clear_workers(self):
        if not self.workplace_var.get():
            messagebox.showerror("Error", "Please select a workplace first")
            return
            
        if not messagebox.askyesno("Confirm", "Are you sure you want to delete ALL workers for this workplace?"):
            return
            
        try:
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            
            # get workplace ID
            c.execute('SELECT id FROM workplaces WHERE name = ?', (self.workplace_var.get(),))
            workplace_id = c.fetchone()[0]
            
            # get worker IDs
            c.execute('SELECT id FROM workers WHERE workplace_id = ?', (workplace_id,))
            worker_ids = c.fetchall()
            
            # delete availability records
            for worker_id in worker_ids:
                c.execute('DELETE FROM availability WHERE worker_id = ?', (worker_id[0],))
                
            # delete workers
            c.execute('DELETE FROM workers WHERE workplace_id = ?', (workplace_id,))
            
            conn.commit()
            conn.close()
            
            self.update_worker_list()
            messagebox.showinfo("Success", "All workers deleted successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete workers: {str(e)}")
    
    def load_shifts(self):
        if not self.shift_workplace_var.get():
            messagebox.showerror("Error", "Please select a workplace first")
            return
            
        try:
            # clear shift list
            for item in self.shift_list.get_children():
                self.shift_list.delete(item)
                
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            
            # get workplace ID
            c.execute('SELECT id FROM workplaces WHERE name = ?', (self.shift_workplace_var.get(),))
            workplace_id = c.fetchone()[0]
            
            # get shifts
            c.execute('''SELECT id, day, start_time, end_time, positions
                        FROM shifts
                        WHERE workplace_id = ?
                        ORDER BY CASE 
                            WHEN day = 'Sunday' THEN 1
                            WHEN day = 'Monday' THEN 2
                            WHEN day = 'Tuesday' THEN 3
                            WHEN day = 'Wednesday' THEN 4
                            WHEN day = 'Thursday' THEN 5
                            WHEN day = 'Friday' THEN 6
                            WHEN day = 'Saturday' THEN 7
                        END''', (workplace_id,))
            shifts = c.fetchall()
            
            conn.close()
            
            for shift in shifts:
                self.shift_list.insert('', 'end', values=shift)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load shifts: {str(e)}")
    
    def add_shift(self):
        if not self.shift_workplace_var.get():
            messagebox.showerror("Error", "Please select a workplace first")
            return
            
        day = self.shift_day_var.get()
        start_time = self.shift_start_time.get()
        end_time = self.shift_end_time.get()
        positions = self.shift_positions.get()
        
        if not all([day, start_time, end_time, positions]):
            messagebox.showerror("Error", "Please fill all fields")
            return
            
        try:
            # validate time format
            try:
                datetime.strptime(start_time, "%I:%M %p")
                datetime.strptime(end_time, "%I:%M %p")
                positions = int(positions)
            except ValueError:
                messagebox.showerror("Error", "Invalid input format. Time must be in format 'HH:MM AM/PM' and positions must be a number.")
                return
                
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            
            # get workplace ID
            c.execute('SELECT id FROM workplaces WHERE name = ?', (self.shift_workplace_var.get(),))
            workplace_id = c.fetchone()[0]
            
            # add shift
            c.execute('''INSERT INTO shifts 
                        (workplace_id, day, start_time, end_time, positions)
                        VALUES (?, ?, ?, ?, ?)''',
                     (workplace_id, day, start_time, end_time, positions))
            
            conn.commit()
            conn.close()
            
            self.load_shifts()
            messagebox.showinfo("Success", "Shift added successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add shift: {str(e)}")
    
    def delete_shift(self):
        selected = self.shift_list.selection()
        if not selected:
            messagebox.showerror("Error", "Please select a shift to delete")
            return
            
        shift_id = self.shift_list.item(selected[0], "values")[0]
        
        try:
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            
            c.execute('DELETE FROM shifts WHERE id = ?', (shift_id,))
            
            conn.commit()
            conn.close()
            
            self.load_shifts()
            messagebox.showinfo("Success", "Shift deleted successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete shift: {str(e)}")
    
    def time_to_datetime(self, time_str, date_str=None):
        # convert "9:00 AM" to datetime object
        time_obj = datetime.strptime(time_str, "%I:%M %p")
        
        if date_str:
            # if date provided, combine with time
            date_obj = datetime.strptime(date_str, "%Y-%m-%d")
            return datetime.combine(date_obj.date(), time_obj.time())
        else:
            return time_obj
    
    def generate_schedule(self):
        if not self.schedule_workplace_var.get():
            messagebox.showerror("Error", "Please select a workplace")
            return
            
        try:
            # get date
            start_date_str = self.start_date_entry.get()
            try:
                start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("Error", "Invalid date format. Use YYYY-MM-DD")
                return
                
            # get workplace info
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            
            # get workplace hours and ID
            c.execute('''SELECT id, hours_open, hours_close 
                        FROM workplaces 
                        WHERE name = ?''', (self.schedule_workplace_var.get(),))
            workplace = c.fetchone()
            
            if not workplace:
                messagebox.showerror("Error", "Workplace not found")
                conn.close()
                return
                
            workplace_id = workplace[0]
            
            # get shifts for this workplace
            c.execute('''SELECT day, start_time, end_time, positions
                        FROM shifts
                        WHERE workplace_id = ?''', (workplace_id,))
            shifts_data = c.fetchall()
            
            # if no shifts defined, create default shifts based on workplace hours
            if not shifts_data:
                messagebox.showinfo("Info", "No shifts defined. Using workplace hours to create a single shift per day.")
                default_shifts = {}
                for day in ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']:
                    default_shifts[day] = [(workplace[1], workplace[2], 1)]  # (start, end, positions)
            else:
                # group shifts by day
                default_shifts = {}
                for day, start, end, positions in shifts_data:
                    if day not in default_shifts:
                        default_shifts[day] = []
                    default_shifts[day].append((start, end, positions))
            
            # get workers and their availability
            c.execute('''SELECT w.id, w.first_name, w.last_name, w.work_study,
                               a.day, a.start_time, a.end_time
                        FROM workers w
                        JOIN availability a ON w.id = a.worker_id
                        WHERE w.workplace_id = ?''', (workplace_id,))
            availability_data = c.fetchall()
            
            # group availability by worker and day
            worker_availability = {}
            for worker_id, fname, lname, work_study, day, start, end in availability_data:
                if worker_id not in worker_availability:
                    worker_availability[worker_id] = {
                        'name': f"{fname} {lname}",
                        'work_study': work_study,
                        'days': {}
                    }
                
                if day not in worker_availability[worker_id]['days']:
                    worker_availability[worker_id]['days'][day] = []
                    
                worker_availability[worker_id]['days'][day].append((start, end))
            
            conn.close()
            
            # generate schedule for a week
            schedule = self.create_weekly_schedule(worker_availability, default_shifts, start_date)
            
            # display schedule
            self.display_schedule(schedule)
            
            messagebox.showinfo("Success", "Schedule generated successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate schedule: {str(e)}")
    
    def create_weekly_schedule(self, worker_availability, shifts_by_day, start_date):
        # create empty schedule
        days_of_week = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
        current_date = start_date
        schedule = {}
        
        # for each day of the week
        for day_offset in range(7):
            day_name = days_of_week[current_date.weekday() if current_date.weekday() < 7 else 0]
            day_date = current_date.strftime("%Y-%m-%d")
            
            # skip days with no shifts
            if day_name not in shifts_by_day:
                current_date += timedelta(days=1)
                continue
                
            # process each shift for this day
            for shift_start, shift_end, positions in shifts_by_day[day_name]:
                shift_key = f"{shift_start} - {shift_end}"
                
                if shift_key not in schedule:
                    schedule[shift_key] = {day: [] for day in days_of_week}
                
                # find available workers for this shift
                available_workers = []
                for worker_id, worker_data in worker_availability.items():
                    if day_name in worker_data['days']:
                        # check if worker is available during this shift
                        for avail_start, avail_end in worker_data['days'][day_name]:
                            shift_start_time = self.time_to_datetime(shift_start)
                            shift_end_time = self.time_to_datetime(shift_end)
                            avail_start_time = self.time_to_datetime(avail_start)
                            avail_end_time = self.time_to_datetime(avail_end)
                            
                            # worker is available if their availability covers the entire shift
                            if avail_start_time <= shift_start_time and avail_end_time >= shift_end_time:
                                available_workers.append({
                                    'id': worker_id,
                                    'name': worker_data['name'],
                                    'work_study': worker_data['work_study']
                                })
                                break
                
                # sort workers by work study status (prioritize work study)
                available_workers.sort(key=lambda x: x['work_study'], reverse=True)
                
                # assign workers to positions
                assigned_workers = available_workers[:positions]
                schedule[shift_key][day_name] = [worker['name'] for worker in assigned_workers]
            
            current_date += timedelta(days=1)
        
        return schedule
    
    def display_schedule(self, schedule):
        # clear existing items
        for item in self.schedule_display.get_children():
            self.schedule_display.delete(item)
            
        # sort shifts by start time
        sorted_shifts = sorted(schedule.keys(), key=lambda x: self.time_to_datetime(x.split(' - ')[0]))
        
        # display schedule
        days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
        
        for shift in sorted_shifts:
            row_data = [shift]  # time column
            
            for day in days:
                workers = schedule[shift][day]
                if workers:
                    row_data.append('\n'.join(workers))
                else:
                    row_data.append('')
            
            self.schedule_display.insert('', 'end', values=tuple(row_data))
    
    def export_schedule(self):
        if not self.schedule_display.get_children():
            messagebox.showerror("Error", "No schedule to export. Please generate a schedule first.")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not filename:
            return
            
        try:
            # create DataFrame from treeview
            data = []
            columns = ["Time", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
            
            for item_id in self.schedule_display.get_children():
                values = self.schedule_display.item(item_id, 'values')
                data.append(values)
            
            df = pd.DataFrame(data, columns=columns)
            
            # export to excel
            df.to_excel(filename, index=False)
            
            messagebox.showinfo("Success", f"Schedule exported to {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export schedule: {str(e)}")

def main():
    root = tk.Tk()
    app = SchedulerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
