# MRX_3.2

MRX_3.2 is a Python-based desktop application built using `Tkinter` for the GUI.
It allows users to upload a CSV file, process the data, and generate customized Excel reports.
The application includes responsive UI elements, progress bars, hover effects, and an automated Excel styling feature using OpenPyXL. 

# Download Link -: https://drive.google.com/drive/folders/1txzkwpj2auxUB5T3sxJwIerNk91cdZG2


## Introduction

The mrx_3.2 application was created as part of my journey to develop advanced Python skills. The project focuses on uploading and processing CSV files related to Consumer Information (CI) reports and generating Excel files for specific sub-divisions. The application also has a professional UI with hover effects and interactive components like progress bars and notifications.

## Features

- CSV Upload: Users can upload a CSV file that gets processed and filtered according to specific criteria.
- Data Cleaning & Filtering: Unwanted columns are removed from the data, and only pending surveys from specific sub-divisions are selected.
- Responsive Progress Bar: The application visually indicates the progress of data processing.
- Excel Report Generation: Users can generate and save filtered Excel reports for two sub-divisions, styled using OpenPyXL.
- Countdown Timer: A countdown effect appears when saving reports, enhancing user engagement.
- LinkedIn & Instagram Links: Quick access buttons to visit my LinkedIn and Instagram profiles.
- Gradient Background & Hover Effects: The GUI has an attractive gradient background with interactive button hover effects.

## Libraries Used

Here are the key Python libraries utilized in this project:

- Tkinter: For building the graphical user interface (GUI).
- Pandas: For handling data manipulation and CSV processing.
- OpenPyXL: To work with Excel files, including styling, formatting, and saving.
- Threading: For running tasks concurrently (e.g., Excel styling and saving reports).
- Time: For simulating delays (e.g., countdown timer) and updating the UI efficiently.
- Webbrowser: To open external links (LinkedIn and Instagram profiles).

## How It Works

1. CSV File Upload: The user uploads a CI report in CSV format. The system processes the CSV file and removes irrelevant columns.
2. Data Processing: It filters the data for pending resurvey statuses, and sorts it by 'Survey done by.'
3. Excel File Generation: Based on the sub-division, two Excel files are generated (`Sub_Div-1_Resurvey.xlsx` and `Sub_Div-2_Resurvey.xlsx`).
4. Excel Styling: The application uses OpenPyXL to apply a light blue header style and adjust column widths for better readability.
5. Saving Reports: After processing, the reports can be saved in a user-specified folder with notifications indicating successful completion.
6. User Interface Enhancements: Includes a progress bar, countdown timer for saving, hover effects on buttons, and social media links.

## What I Learned

During the development of **mrx_3.2**, I learned and improved my understanding of the following concepts:

1. GUI Development with Tkinter: Creating and structuring a responsive and visually appealing user interface.
2. Multithreading: Implementing threading to manage time-consuming tasks like styling and saving Excel files without freezing the GUI.
3. Data Manipulation with Pandas: Efficiently processing large CSV files and filtering data based on specific criteria.
4. Excel Automation with OpenPyXL: Automating Excel file generation and applying advanced styles, such as adjusting column width and coloring headers.
5. Progress Indicators: Using Tkinter progress bars and labels to visually display processing progress to users.
6. File Dialog and Saving Mechanism: Implementing file dialogs for selecting CSV files and choosing directories to save output files.
7. Hover Effects: Enhancing UI elements with hover effects to create a more interactive user experience.
8. Web Automation: Using the webbrowser module to integrate external links (LinkedIn and Instagram) for user engagement.

## How to Use

1. Clone this repository:
   
2. Install the required dependencies:
  
3. Run the `app.py` script:
   
4. Upload your `CI_Report.csv` file and follow the on-screen instructions to generate and save the reports.

## Future Improvements

- PDF Generation: Add functionality to export the generated Excel files to PDF format.
- Progress Circle: Enhance the user experience with a progress circle, replacing the progress bar.
- More File Formats: Allow uploading and processing other formats like Excel files (`.xlsx`).
- Error Logging: Add comprehensive logging to better handle and troubleshoot errors.

## Credits

Developed by Soumy Chauhan https://www.linkedin.com/in/soumy-chauhan/ All rights reserved © 2024 Soumy Chauhan.


![MRX_3 2](https://github.com/user-attachments/assets/418b0b37-c944-4892-ab36-54e3fa611c29)
