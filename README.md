# <p align="center">Excel and Twitter Data Integration API</p>

<p align="center">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium&logoColor=white" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"></a>
</p>

## Introduction

This project provides an API to integrate Excel data manipulation with real-time Twitter data analysis. It allows users to perform queries against Excel data, automate Excel tasks, and monitor Twitter for specific events, such as emergencies in particular locations. The target users are data analysts, automation engineers, and anyone looking to combine social media insights with structured data.

## Table of Contents

1.  [Key Features](#key-features)
2.  [Installation Guide](#installation-guide)
3.  [Usage](#usage)
4.  [Environment Variables](#environment-variables)
5.  [Project Structure](#project-structure)
6.  [Technologies Used](#technologies-used)
7.  [License](#license)

## Key Features

*   **Excel Data Querying:** Process natural language queries to retrieve and manipulate data within Excel spreadsheets.
*   **Excel Task Automation:** Automate repetitive tasks in Excel based on user commands.
*   **Real-time Twitter Monitoring:** Scrape and analyze tweets based on keywords and location to identify relevant events.
*   **Emergency Event Detection:** Identify tweets related to emergencies in specific locations using keyword-based analysis.
*   **Health Check Endpoint:** API endpoint to check the status of the API, Excel connection, and other dependencies.
*   **Modular Design:** Utilizes an `excel_agent` object for managing Excel-related functionalities, promoting code reusability and maintainability.

## Installation Guide

1.  **Clone the repository:**

    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2.  **Install dependencies:**

    ```bash
    pip install -r requirements.txt
    ```

    *Note: A `requirements.txt` file was not provided but is required for this step.  Include the following packages at minimum: `fastapi`, `uvicorn`, `selenium`, `pandas`, `webdriver_manager`.*

3.  **Set up environment variables:**

    Create a `.env` file in the project root and populate it with the necessary environment variables (see [Environment Variables](#environment-variables) section).

4.  **Run the FastAPI server:**

    ```bash
    uvicorn app:app --reload
    ```

## Usage

The API provides several endpoints for interacting with Excel and Twitter data.

*   **/index**: Renders the main page (`index.html`).
*   **/process_query**: Accepts a query related to Excel data and returns the processed result.
*   **/autonomous_action**: Accepts a query, context, and image data, processes it and returns the result.  This likely enables AI-driven actions in Excel.
*   **/connect_to_excel**: Establishes a connection to the Excel application.
*   **/health_check**: Returns the health status of the API and the Excel connection.
*   **/execute_command**: Executes a specific command within Excel.

Example (replace `your_query`):

```bash
curl -X POST -H "Content-Type: application/json" -d '{"query": "your_query"}' http://localhost:8000/process_query
```

## Environment Variables

The following environment variables are required:

*   `EXCEL_CONNECTION_STRING`: The connection string to your Excel instance. *Note: This is a placeholder, actual implementation may vary.*
*   `TWITTER_SEARCH_QUERY`: The search query used for scraping tweets.
*   `CHROME_PROFILE_PATH`: Path to the Chrome profile.
*   `CHROME_PROFILE_NAME`: Name of the Chrome profile.
*   `EMERGENCY_KEYWORDS`: Comma-separated keywords to identify emergency-related tweets.
*   `LOCATIONS`: Comma-separated list of locations to filter tweets.

*NOTE: add any other variables that are crucial to the program and that are not defined here*

## Project Structure

```
.
├── app.py       # FastAPI application for handling API requests
├── main.py      # Web scraping functions for Twitter
└── README.md    # Project documentation
```

## Technologies Used

<p align="left">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium&logoColor=white" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"></a>
</p>

*   **Backend:** FastAPI (Python)
*   **Web Scraping:** Selenium
*   **Data Analysis:** Pandas
*   **Web Driver Management:** `webdriver_manager`

## License

MIT License