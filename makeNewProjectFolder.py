import os
import datetime

# Prepare time object
today = datetime.date.today()

# Make path
newpath = today.strftime('C:\- Projects\%Y-%m-%d - Description')

# Check existing, then create main directory
if not os.path.exists(newpath): os.makedirs(newpath)

# Make Subdirectory
os.makedirs(newpath + "\Exports")

