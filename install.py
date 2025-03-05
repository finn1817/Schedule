import subprocess
import sys
import os
import sqlite3

def install_requirements():
    try:
        # install all required packages
        requirements = ['pandas', 'openpyxl']  # openpyxl for Excel
        for package in requirements:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        
        print("All requirements installed successfully!")
        
        # make data directory
        if not os.path.exists('data'):
            os.makedirs('data')
            
        # make database
        conn = sqlite3.connect('data/schedule.db')
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
        
        conn.commit()
        conn.close()
        
        print("Database created successfully!")
        print("Installation complete! You can now run scheduler.py")
        
    except Exception as e:
        print(f"An error occurred during installation: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    install_requirements()
