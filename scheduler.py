import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime, timedelta
import sqlite3
import re

class SchedulerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Work Schedule Manager")
        self.root.geometry("1200x800")
        
        # start the database connection
        self.db_file = 'data/schedule.db'
        
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
        
    def setup_workplace_tab(self):
        workplace_frame = ttk.Frame(self.notebook)
        self.notebook.add(workplace_frame, text="Workplace Management")
        
        # workplace controls
        ttk.Label(workplace_frame, text="Workplace Name:").grid(row=0, column=0, pady=5)
        self.workplace_name = ttk.Entry(workplace_frame)
        self.workplace_name.grid(row=0, column=1, pady=5)
        
        ttk.Label(workplace_frame, text="Opening Time (HH:MM):").grid(row=1, column=0, pady=5)
        self.opening_time = ttk.Entry(workplace_frame)
        self.opening_time.grid(row=1, column=1, pady=5)
        
        ttk.Label(workplace_frame, text="Closing Time (HH:MM):").grid(row=2, column=0, pady=5)
        self.closing_time = ttk.Entry(workplace_frame)
        self.closing_time.grid(row=2, column=1, pady=5)
        
        ttk.Button(workplace_frame, text="Add Workplace", 
                  command=self.add_workplace).grid(row=3, column=0, columnspan=2, pady=10)
        
        # workplace list
        self.workplace_list = ttk.Treeview(workplace_frame, columns=("Name", "Hours"))
        self.workplace_list.grid(row=4, column=0, columnspan=2, pady=10)
        self.workplace_list.heading("Name", text="Name")
        self.workplace_list.heading("Hours", text="Operating Hours")
        
        self.update_workplace_list()
        
    def setup_import_tab(self):
        import_frame = ttk.Frame(self.notebook)
        self.notebook.add(import_frame, text="Import Workers")
        
        # workplace selection
        ttk.Label(import_frame, text="Select Workplace:").pack(pady=5)
        self.workplace_var = tk.StringVar()
        self.workplace_dropdown = ttk.Combobox(import_frame, textvariable=self.workplace_var)
        self.workplace_dropdown.pack(pady=5)
        
        # excel import
        ttk.Button(import_frame, text="Import Excel File", 
                  command=self.import_excel).pack(pady=10)
        
        # worker list
        self.worker_list = ttk.Treeview(import_frame, 
                                      columns=("Name", "Email", "Work Study"))
        self.worker_list.pack(pady=10, fill=tk.BOTH, expand=True)
        self.worker_list.heading("Name", text="Name")
        self.worker_list.heading("Email", text="Email")
        self.worker_list.heading("Work Study", text="Work Study")
        
    def setup_schedule_tab(self):
        schedule_frame = ttk.Frame(self.notebook)
        self.notebook.add(schedule_frame, text="Generate Schedule")
        
        # controls frame
        controls = ttk.Frame(schedule_frame)
        controls.pack(fill=tk.X, pady=10)
        
        ttk.Label(controls, text="Select Workplace:").pack(side=tk.LEFT, padx=5)
        self.schedule_workplace_var = tk.StringVar()
        self.schedule_workplace_dropdown = ttk.Combobox(controls, 
                                                      textvariable=self.schedule_workplace_var)
        self.schedule_workplace_dropdown.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(controls, text="Generate Schedule", 
                  command=self.generate_schedule).pack(side=tk.LEFT, padx=5)
        
        # schedule display
        self.schedule_display = ttk.Treeview(schedule_frame, 
                                           columns=("Time", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"))
        self.schedule_display.pack(fill=tk.BOTH, expand=True, pady=10)
        
        for col in ("Time", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"):
            self.schedule_display.heading(col, text=col)
    
    def add_workplace(self):
        name = self.workplace_name.get()
        open_time = self.opening_time.get()
        close_time = self.closing_time.get()
        
        if not all([name, open_time, close_time]):
            messagebox.showerror("Error", "Please fill all fields")
            return
            
        try:
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            c.execute('''INSERT INTO workplaces (name, hours_open, hours_close)
                        VALUES (?, ?, ?)''', (name, open_time, close_time))
            conn.commit()
            conn.close()
            
            self.update_workplace_list()
            self.workplace_name.delete(0, tk.END)
            self.opening_time.delete(0, tk.END)
            self.closing_time.delete(0, tk.END)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add workplace: {str(e)}")
    
    def update_workplace_list(self):
        # clearing existing items
        for item in self.workplace_list.get_children():
            self.workplace_list.delete(item)
            
        # fetch and display workplaces
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('SELECT name, hours_open, hours_close FROM workplaces')
        workplaces = c.fetchall()
        conn.close()
        
        for workplace in workplaces:
            self.workplace_list.insert('', 'end', values=(workplace[0], 
                                                        f"{workplace[1]} - {workplace[2]}"))
            
        # update the workplace dropdowns
        workplace_names = [w[0] for w in workplaces]
        self.workplace_dropdown['values'] = workplace_names
        self.schedule_workplace_dropdown['values'] = workplace_names
    
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
            for _, row in df.iterrows():
                # adding workers
                c.execute('''INSERT INTO workers 
                            (workplace_id, first_name, last_name, email, work_study)
                            VALUES (?, ?, ?, ?, ?)''',
                         (workplace_id, row['First Name'], row['Last Name'], 
                          row['Email'], row['Work Study'] == 'Y'))
                worker_id = c.lastrowid
                
                # adding availability
                for day in ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 
                           'Thursday', 'Friday', 'Saturday']:
                    if pd.notna(row[day]) and row[day].lower() != 'na':
                        start, end = self.parse_time_range(row[day])
                        c.execute('''INSERT INTO availability 
                                    (worker_id, day, start_time, end_time)
                                    VALUES (?, ?, ?, ?)''',
                                 (worker_id, day, start, end))
            
            conn.commit()
            conn.close()
            
            self.update_worker_list()
            messagebox.showinfo("Success", "Workers imported successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import Excel file: {str(e)}")
    
    def parse_time_range(self, time_str):
        # converting "2 pm - 12 am" format to 24-hour times
        times = time_str.lower().split('-')
        start = datetime.strptime(times[0].strip(), '%I %p').strftime('%H:%M')
        end = datetime.strptime(times[1].strip(), '%I %p').strftime('%H:%M')
        return start, end
    
    def update_worker_list(self):
        # clearing all existing items
        for item in self.worker_list.get_children():
            self.worker_list.delete(item)
            
        if not self.workplace_var.get():
            return
            
        # fetching and displaying workers
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('''SELECT first_name, last_name, email, work_study 
                    FROM workers 
                    JOIN workplaces ON workers.workplace_id = workplaces.id 
                    WHERE workplaces.name = ?''', (self.workplace_var.get(),))
        workers = c.fetchall()
        conn.close()
        
        for worker in workers:
            self.worker_list.insert('', 'end', 
                                  values=(f"{worker[0]} {worker[1]}", 
                                        worker[2], 
                                        'Yes' if worker[3] else 'No'))
    
    def generate_schedule(self):
        if not self.schedule_workplace_var.get():
            messagebox.showerror("Error", "Please select a workplace")
            return
            
        try:
            # getting workplace info
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            
            # getting workplace hours
            c.execute('''SELECT id, hours_open, hours_close 
                        FROM workplaces 
                        WHERE name = ?''', (self.schedule_workplace_var.get(),))
            workplace = c.fetchone()
            workplace_id = workplace[0]
            
            # getting workers and their availability
            c.execute('''SELECT w.id, w.first_name, w.last_name, w.work_study,
                               a.day, a.start_time, a.end_time
                        FROM workers w
                        JOIN availability a ON w.id = a.worker_id
                        WHERE w.workplace_id = ?''', (workplace_id,))
            availability_data = c.fetchall()
            
            conn.close()
            
            # make (generate) schedule
            schedule = self.create_weekly_schedule(availability_data, workplace)
            
            # show (display) schedule
            self.display_schedule(schedule)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate schedule: {str(e)}")
    
    def create_weekly_schedule(self, availability_data, workplace):
        schedule = {day: {} for day in ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 
                                      'Thursday', 'Friday', 'Saturday']}
        
        # group availability by day
        availability_by_day = {}
        for worker_id, fname, lname, work_study, day, start, end in availability_data:
            if day not in availability_by_day:
                availability_by_day[day] = []
            availability_by_day[day].append({
                'worker_id': worker_id,
                'name': f"{fname} {lname}",
                'work_study': work_study,
                'start': start,
                'end': end
            })
        
        # make a schedule for each day
        for day in schedule:
            if day in availability_by_day:
                workers = availability_by_day[day]
                
                # sort the workers by work study status (prioritize work study students)
                workers.sort(key=lambda x: x['work_study'], reverse=True)
                
                # assign shifts
                current_time = datetime.strptime(workplace[1], '%H:%M')
                end_time = datetime.strptime(workplace[2], '%H:%M')
                
                while current_time < end_time:
                    # find available worker
                    for worker in workers:
                        worker_start = datetime.strptime(worker['start'], '%H:%M')
                        worker_end = datetime.strptime(worker['end'], '%H:%M')
                        
                        if worker_start <= current_time and worker_end >= current_time:
                            # assign 3-hour shift or until worker's end time
                            shift_end = min(
                                current_time + timedelta(hours=3),
                                worker_end,
                                end_time
                            )
                            
                            time_slot = current_time.strftime('%H:%M')
                            if time_slot not in schedule[day]:
                                schedule[day][time_slot] = worker['name']
                            
                            current_time = shift_end
                            break
                    else:
                        # no available worker found, move to next hour
                        current_time += timedelta(hours=1)
        
        return schedule
    
    def display_schedule(self, schedule):
        # clear existing items
        for item in self.schedule_display.get_children():
            self.schedule_display.delete(item)
            
        # get all unique time slots
        time_slots = set()
        for day in schedule.values():
            time_slots.update(day.keys())
        time_slots = sorted(time_slots)
        
        # display schedule
        for time_slot in time_slots:
            row_data = [time_slot]
            for day in ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 
                       'Thursday', 'Friday', 'Saturday']:
                row_data.append(schedule[day].get(time_slot, ''))
            
            self.schedule_display.insert('', 'end', values=tuple(row_data))

def main():
    root = tk.Tk()
    app = SchedulerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
