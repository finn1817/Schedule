import tkinter as tk
from tkinter import ttk, filedialog, messagebox
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
        
        # make main notebook
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # make tabs
        self.setup_workplace_tab()
        self.setup_import_tab()
        self.setup_schedule_tab()
        
        # update workplace lists
        self.load_workplaces()

    def ensure_database_exists(self):
        if not os.path.exists('data'):
            os.makedirs('data')
            
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        
        # make workplaces table with operating hours for each day
        c.execute('''CREATE TABLE IF NOT EXISTS workplaces
                    (id INTEGER PRIMARY KEY,
                     name TEXT NOT NULL,
                     sunday_open TEXT,
                     sunday_close TEXT,
                     monday_open TEXT,
                     monday_close TEXT,
                     tuesday_open TEXT,
                     tuesday_close TEXT,
                     wednesday_open TEXT,
                     wednesday_close TEXT,
                     thursday_open TEXT,
                     thursday_close TEXT,
                     friday_open TEXT,
                     friday_close TEXT,
                     saturday_open TEXT,
                     saturday_close TEXT)''')
        
        # make workers table
        c.execute('''CREATE TABLE IF NOT EXISTS workers
                    (id INTEGER PRIMARY KEY,
                     workplace_id INTEGER,
                     first_name TEXT NOT NULL,
                     last_name TEXT NOT NULL,
                     email TEXT NOT NULL,
                     work_study BOOLEAN NOT NULL,
                     FOREIGN KEY (workplace_id) REFERENCES workplaces(id))''')
        
        # make availability table
        c.execute('''CREATE TABLE IF NOT EXISTS availability
                    (id INTEGER PRIMARY KEY,
                     worker_id INTEGER,
                     day TEXT NOT NULL,
                     start_time TEXT NOT NULL,
                     end_time TEXT NOT NULL,
                     FOREIGN KEY (worker_id) REFERENCES workers(id))''')
        
        # make the shifts table
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
        self.notebook.add(workplace_frame, text="Workplace Hours")
        
        # workplace name input
        ttk.Label(workplace_frame, text="Workplace Name:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.workplace_name = ttk.Entry(workplace_frame, width=30)
        self.workplace_name.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        # create frame for hours
        hours_frame = ttk.LabelFrame(workplace_frame, text="Operating Hours (HH:MM AM/PM)")
        hours_frame.grid(row=1, column=0, columnspan=4, padx=10, pady=10, sticky='nsew')
        
        # day headers
        days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        ttk.Label(hours_frame, text="Day").grid(row=0, column=0, padx=5, pady=5)
        ttk.Label(hours_frame, text="Open").grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(hours_frame, text="Close").grid(row=0, column=2, padx=5, pady=5)
        
        # operating hours inputs for each day
        self.hours_entries = {}
        for i, day in enumerate(days):
            ttk.Label(hours_frame, text=day).grid(row=i+1, column=0, padx=5, pady=5, sticky='w')
            
            # open time entry
            open_entry = ttk.Entry(hours_frame, width=15)
            open_entry.grid(row=i+1, column=1, padx=5, pady=5)
            open_entry.insert(0, "12:00 PM")
            
            # close time entry
            close_entry = ttk.Entry(hours_frame, width=15)
            close_entry.grid(row=i+1, column=2, padx=5, pady=5)
            close_entry.insert(0, "12:00 AM")
            
            self.hours_entries[day] = (open_entry, close_entry)
        
        # buttons
        buttons_frame = ttk.Frame(workplace_frame)
        buttons_frame.grid(row=2, column=0, columnspan=4, padx=10, pady=10, sticky='w')
        
        ttk.Button(buttons_frame, text="Add/Update Workplace", 
                  command=self.save_workplace).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(buttons_frame, text="Load Workplace", 
                  command=self.load_workplace_hours).pack(side=tk.LEFT, padx=5)
        
        # workplace selection dropdown
        ttk.Label(buttons_frame, text="Select Workplace:").pack(side=tk.LEFT, padx=5)
        self.workplace_dropdown_var = tk.StringVar()
        self.workplace_dropdown = ttk.Combobox(buttons_frame, 
                                             textvariable=self.workplace_dropdown_var, 
                                             width=30)
        self.workplace_dropdown.pack(side=tk.LEFT, padx=5)
        
        # shift management section
        shift_frame = ttk.LabelFrame(workplace_frame, text="Shift Management")
        shift_frame.grid(row=3, column=0, columnspan=4, padx=10, pady=10, sticky='nsew')
        
        # shift inputs
        ttk.Label(shift_frame, text="Day:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.shift_day_var = tk.StringVar()
        self.shift_day_dropdown = ttk.Combobox(shift_frame, 
                                             textvariable=self.shift_day_var,
                                             values=days, 
                                             width=15)
        self.shift_day_dropdown.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        ttk.Label(shift_frame, text="Start Time:").grid(row=0, column=2, padx=5, pady=5, sticky='w')
        self.shift_start = ttk.Entry(shift_frame, width=15)
        self.shift_start.grid(row=0, column=3, padx=5, pady=5, sticky='w')
        self.shift_start.insert(0, "12:00 PM")
        
        ttk.Label(shift_frame, text="End Time:").grid(row=0, column=4, padx=5, pady=5, sticky='w')
        self.shift_end = ttk.Entry(shift_frame, width=15)
        self.shift_end.grid(row=0, column=5, padx=5, pady=5, sticky='w')
        self.shift_end.insert(0, "8:00 PM")
        
        ttk.Label(shift_frame, text="Positions:").grid(row=0, column=6, padx=5, pady=5, sticky='w')
        self.shift_positions = ttk.Spinbox(shift_frame, from_=1, to=10, width=5)
        self.shift_positions.grid(row=0, column=7, padx=5, pady=5, sticky='w')
        
        ttk.Button(shift_frame, text="Add Shift", 
                  command=self.add_shift).grid(row=0, column=8, padx=5, pady=5)
        
        # shift list
        self.shift_list = ttk.Treeview(shift_frame, 
                                     columns=("id", "Day", "Start", "End", "Positions"),
                                     show="headings",
                                     height=10)
        self.shift_list.grid(row=1, column=0, columnspan=9, padx=5, pady=5, sticky='nsew')
        
        self.shift_list.heading("id", text="ID")
        self.shift_list.heading("Day", text="Day")
        self.shift_list.heading("Start", text="Start Time")
        self.shift_list.heading("End", text="End Time")
        self.shift_list.heading("Positions", text="Positions")
        
        self.shift_list.column("id", width=50)
        
        ttk.Button(shift_frame, text="Delete Selected Shift", 
                  command=self.delete_shift).grid(row=2, column=0, padx=5, pady=5, sticky='w')
        
        # make frames expandable
        workplace_frame.columnconfigure(0, weight=1)
        workplace_frame.columnconfigure(1, weight=1)
        workplace_frame.columnconfigure(2, weight=1)
        workplace_frame.columnconfigure(3, weight=1)
        workplace_frame.rowconfigure(3, weight=1)
        
        shift_frame.columnconfigure(8, weight=1)
        shift_frame.rowconfigure(1, weight=1)
        
    def setup_import_tab(self):
        import_frame = ttk.Frame(self.notebook)
        self.notebook.add(import_frame, text="Import Workers")
        
        # control panel
        control_frame = ttk.Frame(import_frame)
        control_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(control_frame, text="Select Workplace:").pack(side=tk.LEFT, padx=5)
        self.import_workplace_var = tk.StringVar()
        self.import_workplace_dropdown = ttk.Combobox(control_frame, 
                                                    textvariable=self.import_workplace_var,
                                                    width=30)
        self.import_workplace_dropdown.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="Import Excel File", 
                  command=self.import_excel).pack(side=tk.LEFT, padx=5)
                  
        ttk.Button(control_frame, text="View Workers", 
                  command=self.view_workers).pack(side=tk.LEFT, padx=5)
        
        # worker display frame
        worker_frame = ttk.Frame(import_frame)
        worker_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # create worker list
        self.worker_list = ttk.Treeview(worker_frame, 
                                      columns=("id", "Name", "Email", "Work Study", "Availability"),
                                      show="headings")
        self.worker_list.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.worker_list.heading("id", text="ID")
        self.worker_list.heading("Name", text="Name")
        self.worker_list.heading("Email", text="Email")
        self.worker_list.heading("Work Study", text="Work Study")
        self.worker_list.heading("Availability", text="Availability")
        
        self.worker_list.column("id", width=50)
        self.worker_list.column("Name", width=150)
        self.worker_list.column("Email", width=200)
        self.worker_list.column("Work Study", width=100)
        self.worker_list.column("Availability", width=300)
        
    def setup_schedule_tab(self):
        schedule_frame = ttk.Frame(self.notebook)
        self.notebook.add(schedule_frame, text="Generate Schedule")
        
        # control panel
        control_frame = ttk.Frame(schedule_frame)
        control_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(control_frame, text="Select Workplace:").pack(side=tk.LEFT, padx=5)
        self.schedule_workplace_var = tk.StringVar()
        self.schedule_workplace_dropdown = ttk.Combobox(control_frame, 
                                                      textvariable=self.schedule_workplace_var,
                                                      width=30)
        self.schedule_workplace_dropdown.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(control_frame, text="Start Date (YYYY-MM-DD):").pack(side=tk.LEFT, padx=5)
        self.start_date = ttk.Entry(control_frame, width=15)
        self.start_date.pack(side=tk.LEFT, padx=5)
        # set default to current date
        self.start_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        ttk.Button(control_frame, text="Generate Schedule", 
                  command=self.generate_schedule).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="Export Schedule", 
                  command=self.export_schedule).pack(side=tk.LEFT, padx=5)
        
        # schedule display
        schedule_display_frame = ttk.Frame(schedule_frame)
        schedule_display_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # add scrollbars
        y_scrollbar = ttk.Scrollbar(schedule_display_frame, orient=tk.VERTICAL)
        x_scrollbar = ttk.Scrollbar(schedule_display_frame, orient=tk.HORIZONTAL)
        
        self.schedule_display = ttk.Treeview(schedule_display_frame, 
                                           columns=("Time", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"),
                                           show="headings",
                                           yscrollcommand=y_scrollbar.set,
                                           xscrollcommand=x_scrollbar.set)
        
        y_scrollbar.config(command=self.schedule_display.yview)
        x_scrollbar.config(command=self.schedule_display.xview)
        
        # pack elements
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.schedule_display.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # configure treeview columns
        for col in ("Time", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"):
            self.schedule_display.heading(col, text=col)
            self.schedule_display.column(col, width=150, minwidth=100)
    
    def load_workplaces(self):
        """Update all workplace dropdown lists"""
        try:
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            c.execute('SELECT id, name FROM workplaces')
            workplaces = c.fetchall()
            conn.close()
            
            workplace_names = [wp[1] for wp in workplaces]
            workplace_ids = [wp[0] for wp in workplaces]
            
            # create mapping from name to ID
            self.workplace_id_map = dict(zip(workplace_names, workplace_ids))
            
            # update all dropdowns
            self.workplace_dropdown['values'] = workplace_names
            self.import_workplace_dropdown['values'] = workplace_names
            self.schedule_workplace_dropdown['values'] = workplace_names
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load workplaces: {str(e)}")
    
    def save_workplace(self):
        """Save workplace with operating hours for each day"""
        name = self.workplace_name.get()
        
        if not name:
            messagebox.showerror("Error", "Please enter a workplace name")
            return
            
        # collect hours for each day
        hours_data = {}
        for day, (open_entry, close_entry) in self.hours_entries.items():
            open_time = open_entry.get()
            close_time = close_entry.get()
            
            # validate time format
            try:
                datetime.strptime(open_time, "%I:%M %p")
                datetime.strptime(close_time, "%I:%M %p")
            except ValueError:
                messagebox.showerror("Error", f"Invalid time format for {day}. Use HH:MM AM/PM format.")
                return
                
            day_lower = day.lower()
            hours_data[f"{day_lower}_open"] = open_time
            hours_data[f"{day_lower}_close"] = close_time
        
        try:
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            
            # check if workplace already exists
            c.execute('SELECT id FROM workplaces WHERE name = ?', (name,))
            existing = c.fetchone()
            
            if existing:
                # update existing workplace
                workplace_id = existing[0]
                
                update_query = '''UPDATE workplaces SET 
                                sunday_open = ?, sunday_close = ?,
                                monday_open = ?, monday_close = ?,
                                tuesday_open = ?, tuesday_close = ?,
                                wednesday_open = ?, wednesday_close = ?,
                                thursday_open = ?, thursday_close = ?,
                                friday_open = ?, friday_close = ?,
                                saturday_open = ?, saturday_close = ?
                                WHERE id = ?'''
                
                c.execute(update_query, (
                    hours_data['sunday_open'], hours_data['sunday_close'],
                    hours_data['monday_open'], hours_data['monday_close'],
                    hours_data['tuesday_open'], hours_data['tuesday_close'],
                    hours_data['wednesday_open'], hours_data['wednesday_close'],
                    hours_data['thursday_open'], hours_data['thursday_close'],
                    hours_data['friday_open'], hours_data['friday_close'],
                    hours_data['saturday_open'], hours_data['saturday_close'],
                    workplace_id
                ))
                
                message = f"Workplace '{name}' updated successfully!"
                
            else:
                # insert new workplace
                insert_query = '''INSERT INTO workplaces (
                                name, 
                                sunday_open, sunday_close,
                                monday_open, monday_close,
                                tuesday_open, tuesday_close,
                                wednesday_open, wednesday_close,
                                thursday_open, thursday_close,
                                friday_open, friday_close,
                                saturday_open, saturday_close
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''
                
                c.execute(insert_query, (
                    name,
                    hours_data['sunday_open'], hours_data['sunday_close'],
                    hours_data['monday_open'], hours_data['monday_close'],
                    hours_data['tuesday_open'], hours_data['tuesday_close'],
                    hours_data['wednesday_open'], hours_data['wednesday_close'],
                    hours_data['thursday_open'], hours_data['thursday_close'],
                    hours_data['friday_open'], hours_data['friday_close'],
                    hours_data['saturday_open'], hours_data['saturday_close']
                ))
                
                message = f"Workplace '{name}' added successfully!"
            
            conn.commit()
            conn.close()
            
            self.load_workplaces()
            messagebox.showinfo("Success", message)
            
            # load shifts for the current workplace
            self.workplace_dropdown_var.set(name)
            self.load_shifts()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save workplace: {str(e)}")
    
    def load_workplace_hours(self):
        """Load hours for selected workplace"""
        selected_workplace = self.workplace_dropdown_var.get()
        
        if not selected_workplace:
            messagebox.showerror("Error", "Please select a workplace to load")
            return
            
        try:
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            
            # get workplace data
            c.execute('''SELECT 
                        sunday_open, sunday_close,
                        monday_open, monday_close,
                        tuesday_open, tuesday_close,
                        wednesday_open, wednesday_close,
                        thursday_open, thursday_close,
                        friday_open, friday_close,
                        saturday_open, saturday_close
                        FROM workplaces WHERE name = ?''', (selected_workplace,))
            
            workplace_data = c.fetchone()
            conn.close()
            
            if not workplace_data:
                messagebox.showerror("Error", "Workplace not found")
                return
                
            # update workplace name
            self.workplace_name.delete(0, tk.END)
            self.workplace_name.insert(0, selected_workplace)
            
            # update hours entries
            days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
            for i, day in enumerate(days):
                open_entry, close_entry = self.hours_entries[day]
                
                # clear and set new values
                open_entry.delete(0, tk.END)
                close_entry.delete(0, tk.END)
                
                open_entry.insert(0, workplace_data[i*2])
                close_entry.insert(0, workplace_data[i*2 + 1])
            
            # load shifts for this workplace
            self.load_shifts()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load workplace hours: {str(e)}")
    
    def add_shift(self):
        """Add a shift for the selected workplace"""
        workplace = self.workplace_dropdown_var.get()
        
        if not workplace:
            messagebox.showerror("Error", "Please select a workplace first")
            return
            
        day = self.shift_day_var.get()
        start_time = self.shift_start.get()
        end_time = self.shift_end.get()
        positions = self.shift_positions.get()
        
        if not all([day, start_time, end_time]):
            messagebox.showerror("Error", "Please fill in all shift details")
            return
            
        try:
            # validate time format
            try:
                datetime.strptime(start_time, "%I:%M %p")
                datetime.strptime(end_time, "%I:%M %p")
                positions = int(positions)
            except ValueError:
                messagebox.showerror("Error", "Invalid input format. Time must be in HH:MM AM/PM format")
                return
                
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            
            # get workplace ID
            c.execute('SELECT id FROM workplaces WHERE name = ?', (workplace,))
            workplace_id = c.fetchone()[0]
            
            # insert shift
            c.execute('''INSERT INTO shifts 
                        (workplace_id, day, start_time, end_time, positions)
                        VALUES (?, ?, ?, ?, ?)''',
                     (workplace_id, day, start_time, end_time, positions))
            
            conn.commit()
            conn.close()
            
            # refresh shift list
            self.load_shifts()
            messagebox.showinfo("Success", f"Shift added for {day}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add shift: {str(e)}")
    
    def load_shifts(self):
        """Load shifts for the selected workplace"""
        workplace = self.workplace_dropdown_var.get()
        
        if not workplace:
            return
            
        # clear existing shifts
        for item in self.shift_list.get_children():
            self.shift_list.delete(item)
            
        try:
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            
            # get workplace ID
            c.execute('SELECT id FROM workplaces WHERE name = ?', (workplace,))
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
            
            # display shifts
            for shift in shifts:
                self.shift_list.insert('', 'end', values=shift)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load shifts: {str(e)}")
    
    def delete_shift(self):
        """Delete the selected shift"""
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
            
            # refresh shift list
            self.load_shifts()
            messagebox.showinfo("Success", "Shift deleted successfully")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete shift: {str(e)}")
    
    def import_excel(self):
        """Import workers from Excel file"""
        workplace = self.import_workplace_var.get()
        
        if not workplace:
            messagebox.showerror("Error", "Please select a workplace first")
            return
            
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if not filename:
            return
            
        try:
            # read Excel file
            df = pd.read_excel(filename)
            
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            
            # get workplace ID
            c.execute('SELECT id FROM workplaces WHERE name = ?', (workplace,))
            workplace_id = c.fetchone()[0]
            
            # import workers and availability
            imported_count = 0
            for _, row in df.iterrows():
                try:
                    # check required columns
                    if 'First Name' not in row or 'Last Name' not in row or 'Email' not in row:
                        continue
                        
                    # get worker data
                    first_name = row['First Name']
                    last_name = row['Last Name']
                    email = row['Email']
                    
                    # work study (default to N if not present)
                    work_study = False
                    if 'Work Study' in row and pd.notna(row['Work Study']):
                        work_study = row['Work Study'].upper() == 'Y'
                    
                    # add worker to database
                    c.execute('''INSERT INTO workers 
                                (workplace_id, first_name, last_name, email, work_study)
                                VALUES (?, ?, ?, ?, ?)''',
                             (workplace_id, first_name, last_name, email, work_study))
                    
                    worker_id = c.lastrowid
                    imported_count += 1
                    
                    # process availability for each day
                    days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
                    for day in days:
                        if day in row and pd.notna(row[day]) and str(row[day]).lower() != 'na':
                            try:
                                start_time, end_time = self.parse_time_range(str(row[day]))
                                
                                c.execute('''INSERT INTO availability 
                                            (worker_id, day, start_time, end_time)
                                            VALUES (?, ?, ?, ?)''',
                                         (worker_id, day, start_time, end_time))
                            except Exception as e:
                                print(f"Error parsing time for {first_name} {last_name} on {day}: {str(e)}")
                
                except Exception as e:
                    print(f"Error importing worker: {str(e)}")
                    continue
            
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Success", f"Imported {imported_count} workers")
            self.view_workers()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import Excel file: {str(e)}")
    
    def parse_time_range(self, time_str):
        """Parse time range from format like '2 pm - 12 am' to standard 12-hour format"""
        time_str = time_str.strip().lower()
        
        # split into start and end times
        if '-' in time_str:
            parts = time_str.split('-')
        else:
            parts = re.split(r'\s+to\s+', time_str)
            
        start_part = parts[0].strip()
        end_part = parts[1].strip()
        
        # parse start time
        start_match = re.search(r'(\d+)(?::(\d+))?\s*(am|pm)', start_part)
        if not start_match:
            # try without am/pm
            start_match = re.search(r'(\d+)(?::(\d+))?', start_part)
            if start_match:
                hour = int(start_match.group(1))
                minute = start_match.group(2) or "00"
                ampm = "am" if hour < 12 else "pm"
                start_time = f"{hour}:{minute} {ampm}"
            else:
                raise ValueError(f"Cannot parse start time: {start_part}")
        else:
            hour = int(start_match.group(1))
            minute = start_match.group(2) or "00"
            ampm = start_match.group(3)
            start_time = f"{hour}:{minute} {ampm}"
        
        # parse end time
        end_match = re.search(r'(\d+)(?::(\d+))?\s*(am|pm)', end_part)
        if not end_match:
            # try without am/pm
            end_match = re.search(r'(\d+)(?::(\d+))?', end_part)
            if end_match:
                hour = int(end_match.group(1))
                minute = end_match.group(2) or "00"
                ampm = "am" if hour < 12 else "pm"
                end_time = f"{hour}:{minute} {ampm}"
            else:
                raise ValueError(f"Cannot parse end time: {end_part}")
        else:
            hour = int(end_match.group(1))
            minute = end_match.group(2) or "00"
            ampm = end_match.group(3)
            end_time = f"{hour}:{minute} {ampm}"
        
        # standardize format
        start_dt = datetime.strptime(start_time, "%I:%M %p")
        end_dt = datetime.strptime(end_time, "%I:%M %p")
        
        return start_dt.strftime("%I:%M %p"), end_dt.strftime("%I:%M %p")
    
    def view_workers(self):
        """Display workers for the selected workplace"""
        workplace = self.import_workplace_var.get()
        
        if not workplace:
            messagebox.showerror("Error", "Please select a workplace first")
            return
            
        # clear current worker list
        for item in self.worker_list.get_children():
            self.worker_list.delete(item)
            
        try:
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            
            # get workplace ID
            c.execute('SELECT id FROM workplaces WHERE name = ?', (workplace,))
            workplace_id = c.fetchone()[0]
            
            # get workers
            c.execute('''SELECT w.id, w.first_name, w.last_name, w.email, w.work_study
                        FROM workers w
                        WHERE w.workplace_id = ?''', (workplace_id,))
            
            workers = c.fetchall()
            
            # for each worker, get their availability
            for worker in workers:
                worker_id, first_name, last_name, email, work_study = worker
                
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
                
                # format availability for display
                avail_str = ", ".join([f"{day}: {start}-{end}" for day, start, end in availability])
                
                # add to worker list
                self.worker_list.insert('', 'end', 
                                      values=(worker_id, 
                                            f"{first_name} {last_name}", 
                                            email, 
                                            "Yes" if work_study else "No",
                                            avail_str))
            
            conn.close()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load workers: {str(e)}")
    
    def generate_schedule(self):
        """Generate a weekly work schedule"""
        workplace = self.schedule_workplace_var.get()
        
        if not workplace:
            messagebox.showerror("Error", "Please select a workplace")
            return
            
        try:
            # get start date
            start_date_str = self.start_date.get()
            try:
                start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("Error", "Invalid date format. Use YYYY-MM-DD")
                return
                
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            
            # get workplace ID and operating hours
            c.execute('SELECT id FROM workplaces WHERE name = ?', (workplace,))
            workplace_id = c.fetchone()[0]
            
            # get shifts for this workplace
            c.execute('''SELECT day, start_time, end_time, positions
                        FROM shifts
                        WHERE workplace_id = ?''', (workplace_id,))
            
            shifts_data = c.fetchall()
            
            if not shifts_data:
                messagebox.showerror("Error", "No shifts defined for this workplace. Please add shifts in the Workplace Hours tab.")
                conn.close()
                return
                
            # get workers and their availability
            c.execute('''SELECT w.id, w.first_name, w.last_name, w.work_study,
                               a.day, a.start_time, a.end_time
                        FROM workers w
                        JOIN availability a ON w.id = a.worker_id
                        WHERE w.workplace_id = ?''', (workplace_id,))
                        
            availability_data = c.fetchall()
            
            conn.close()
            
            if not availability_data:
                messagebox.showerror("Error", "No workers with availability found")
                return
                
            # generate schedule
            schedule = self.create_schedule(shifts_data, availability_data, start_date)
            
            # display schedule
            self.display_schedule(schedule)
            
            messagebox.showinfo("Success", "Schedule generated successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate schedule: {str(e)}")
    
    def create_schedule(self, shifts_data, availability_data, start_date):
        """Create a weekly schedule based on shifts and worker availability"""
        # map of day names to day numbers (0=Monday in datetime, but we want Sunday=0)
        day_to_num = {
            'Sunday': 6,
            'Monday': 0,
            'Tuesday': 1,
            'Wednesday': 2,
            'Thursday': 3,
            'Friday': 4,
            'Saturday': 5
        }
        
        num_to_day = {v: k for k, v in day_to_num.items()}
        
        # group shifts by day
        shifts_by_day = {}
        for day, start, end, positions in shifts_data:
            if day not in shifts_by_day:
                shifts_by_day[day] = []
            shifts_by_day[day].append((start, end, positions))
        
        # group worker availability by day
        workers_by_day = {}
        for worker_id, fname, lname, work_study, day, start, end in availability_data:
            if day not in workers_by_day:
                workers_by_day[day] = []
                
            workers_by_day[day].append({
                'id': worker_id,
                'name': f"{fname} {lname}",
                'work_study': work_study,
                'start': start,
                'end': end
            })
        
        # create empty schedule organized by shift time
        schedule = {}
        
        # for each shift
        for day, shifts in shifts_by_day.items():
            # for each shift on this day
            for shift_start, shift_end, positions in shifts:
                shift_key = f"{shift_start} - {shift_end}"
                
                if shift_key not in schedule:
                    schedule[shift_key] = {d: [] for d in ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']}
                
                # find available workers for this shift
                if day in workers_by_day:
                    available_workers = []
                    
                    for worker in workers_by_day[day]:
                        # check if worker is available for this shift
                        worker_start = self.time_to_datetime(worker['start'])
                        worker_end = self.time_to_datetime(worker['end'])
                        
                        shift_start_dt = self.time_to_datetime(shift_start)
                        shift_end_dt = self.time_to_datetime(shift_end)
                        
                        # worker is available if their time covers the shift
                        if worker_start <= shift_start_dt and worker_end >= shift_end_dt:
                            available_workers.append(worker)
                    
                    # prioritize work study students
                    available_workers.sort(key=lambda w: (w['work_study'], w['name']), reverse=True)
                    
                    # assign workers to positions
                    assigned_workers = available_workers[:positions]
                    schedule[shift_key][day] = [w['name'] for w in assigned_workers]
        
        return schedule
    
    def time_to_datetime(self, time_str):
        """Convert time string to datetime object"""
        return datetime.strptime(time_str, "%I:%M %p")
    
    def display_schedule(self, schedule):
        """Display the generated schedule"""
        # clear existing items
        for item in self.schedule_display.get_children():
            self.schedule_display.delete(item)
            
        # sort shifts by start time
        def get_start_time(shift_key):
            start_time = shift_key.split(' - ')[0]
            return self.time_to_datetime(start_time)
            
        sorted_shifts = sorted(schedule.keys(), key=get_start_time)
        
        days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
        
        # add rows for each shift
        for shift in sorted_shifts:
            row_data = [shift]
            
            for day in days:
                workers = schedule[shift][day]
                if workers:
                    row_data.append('\n'.join(workers))
                else:
                    row_data.append('')
            
            self.schedule_display.insert('', 'end', values=tuple(row_data))
    
    def export_schedule(self):
        """Export the schedule to Excel"""
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
            # create dataFrame from treeview
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
