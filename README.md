# Automated Report Generator
This project creates structured reports based on topics or uploaded files using AI-driven content generation. Below are the setup instructions to get the project running.

1.Prerequisites
Python (version 3.11.8)

2.Setup
Navigate to the Project Directory

cd C:\Users\Siddhi Sawant\OneDrive\Desktop\aiReport\report

3.Create and Activate a Virtual Environment

First, remove any previous virtual environment:

bash

# Remove old venv if it exists
rmdir /s /q venv

Then, create a new virtual environment:

bash

# Create new venv
python -m venv venv
Activate the virtual environment:

bash

# Activate venv
.\venv\Scripts\activate
Install Required Packages

4.With the virtual environment activated, install the necessary dependencies:

bash

pip install requests==2.31.0 python-docx==1.0.0



5.Running the Script
After completing the setup, you can run the script to generate reports based on the input files or topics provided.
