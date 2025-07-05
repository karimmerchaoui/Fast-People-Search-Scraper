
![Fast_People_Search_Scraper](https://github.com/user-attachments/assets/adbe2316-4cef-4d5c-a412-6aea5a0e1b7e)

# Overview

The Fast People Search Scraper is a powerful GUI application designed to efficiently extract people search data from fastpeoplesearch.com. This tool processes batches of names and locations, scrapes relevant information, and compiles the results into an organized Excel spreadsheet. With multi-threading capabilities, it can handle large datasets quickly while providing real-time progress tracking.
<strong>This solves the problem</strong> of manually searching for and compiling people search data, automating the entire process and saving time while ensuring accuracy and efficiency.

![0705](https://github.com/user-attachments/assets/3884b092-2dda-4f73-9806-a31929f1114b)

# Who Would Benefit from This Tool

- **Real Estate Professionals**: Quickly find property owners and contact information
- **Marketing Teams**: Build targeted contact lists for campaigns
- **Researchers**: Gather demographic data for analysis
- **Legal Professionals**: Locate individuals for legal processes
- **Private Investigators**: Efficiently track person-of-interest information

# Technologies Used

| Component              | Technology Stack           |
|------------------------|----------------------------|
| Core Language          | Python 3.7+                |
| GUI Framework          | CustomTkinter              |
| Web Scraping           | BeautifulSoup4 + Requests  |
| Proxy Management       | ScraperAPI                 |
| Data Processing        | Regex + OpenPyXL           |
| Parallel Execution     | Concurrent.futures         |


# Features

- **Batch Processing**: Handles 7 simultaneous requests per thread
- **Real-time Progress Tracking**: Visual progress bar and completion counter
- **Color-coded Status Indicators**: Immediate visual feedback on processing status
- **Excel Integration**: Direct import/export without conversion
- **Resume Capability**: Saves results incrementally
- **Cancellation Support**: Gracefully stop ongoing operations
- **Detailed Logging**: Comprehensive activity record

# Project Background

This project was originally developed for **MSV Properties** to streamline their property research and owner identification processes. Created by **Karim Merchaoui**, the tool has been optimized for reliability and performance in professional real estate workflows.


# Usage

### 1. Prepare Your Input File

- Create an Excel file with the following columns:
  - `Name`
  - `City` 
  - `State`

  Example:

  | Name           | City      | State |
  |----------------|-----------|-------|
  | John Doe       | New York  | NY    |
  | Michael Brown  | Chicago   | IL    |

### 2. Launch the Application


### 3. Configure Settings in the GUI

- Select your input Excel file.
- Choose your output folder.
- Click **Start Scraping**.

### 4. Monitor Progress

- Watch the real-time progress bar and log messages in the GUI.

### 5. Review Results

- When scraping completes, open the exported Excel file in your chosen output folder.

---


# Output Data Structure
The generated Excel file contains the following information for each person searched:

## Search Results
| Column | Description |
|--------|-------------|
| Name | Full name used in the search query |
| Current Address | Most recently associated mailing or residential address |
| Phone Numbers | List of phone numbers found for the individual |

- The output includes empty rows between entries for easier readability
- Failed or empty results are marked and recorded with placeholder values
  
![image](https://github.com/user-attachments/assets/9dbe6c2c-015d-4346-abf0-117fc588bee9)


