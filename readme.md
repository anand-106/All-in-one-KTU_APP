# <p align="center">Emergency Data Aggregation and Excel Automation</p>

<p align="center">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas" alt="Pandas"></a>
</p>

## Introduction

This project provides a web application for aggregating emergency-related data from Twitter and automating tasks within Excel. It leverages a Flask API, Selenium for web scraping, and interacts with Excel workbooks to provide insights and automate workflows. This tool is ideal for emergency response teams, data analysts, and anyone needing real-time information and spreadsheet automation.

## Table of Contents

1.  [Key Features](#key-features)
2.  [Installation Guide](#installation-guide)
3.  [Usage](#usage)
4.  [Environment Variables](#environment-variables)
5.  [Project Structure](#project-structure)
6.  [Technologies Used](#technologies-used)
7.  [License](#license)

## Key Features

*   **Automated Twitter Scraping:** Scrapes tweets based on a predefined search query using Selenium.
*   **Emergency Tweet Identification:** Processes scraped tweets to identify emergency-related content with location information.
*   **Excel Automation:** Connects to and interacts with Excel instances via an `excel_agent` (implementation not detailed in summaries).
*   **API Endpoints:** Provides Flask API endpoints for processing user queries, executing commands in Excel, and performing health checks.
*   **Data Analysis:** Uses pandas to process and analyze scraped data, adding relevant columns.
*   **Error Handling:** Comprehensive error handling for API requests and web scraping processes.

## Installation Guide

1.  **Clone the repository:**

    ```bash
    git clone <repository_url>
    cd <repository_name>
    ```

2.  **Install dependencies:**

    ```bash
    pip install flask selenium pandas openpyxl chromedriver_autoinstaller
    ```

3.  **Set up environment variables:**
    Create a `.env` file in the project root and define the necessary environment variables (see [Environment Variables](#environment-variables) section).

4.  **Run the FastAPI server:**

    ```bash
    python main.py
    python app.py
    ```

## Usage

1.  **Access the main page:** Open your web browser and navigate to the address where the Flask application is running (e.g., `http://localhost:5000`).

2.  **Interact with the API:** Use the provided API endpoints (described below) to process queries, execute commands in Excel, or check the health of the application.

3.  **API Endpoints:**
    - `/`: Renders the main application page.
    - `/process_query`: Processes a user query and returns an AI response from the `excel_agent`.
    - `/autonomous_action`: Processes a query and autonomously executes actions in Excel.
    - `/execute_command`: Executes a specified command in Excel.
    - `/health_check`: Checks the health of the API, Excel connection, and Gemini configuration.
    - `/connect_to_excel`: Establishes a connection to an Excel instance.

## Environment Variables

The following environment variables are required for the application to function correctly:

*   `SEARCH_QUERY`: The search query used to scrape tweets from Twitter (e.g., "emergency fire").
*   `CHROME_PROFILE_PATH`: The path to the Chrome profile directory.
*   `CHROME_PROFILE_NAME`: The name of the Chrome profile.
*   `SCROLL_ITERATIONS`: The number of iterations to scroll the Twitter page during scraping.
*   `EMERGENCY_KEYWORDS`: comma seperated keywords to define an emergency tweet.
*   `LOCATIONS`: comma seperated locations for location analysis of tweets

## Project Structure

```
.
├── app.py       # Flask API endpoints and application logic
├── main.py      # Twitter scraping and data processing functions
└── readme.md    # Project documentation
```

## Technologies Used

<p align="left">
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Flask-000000?style=for-the-badge&logo=flask&logoColor=white" alt="Flask"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas" alt="Pandas"></a>
  <a href="#"><img src="https://img.shields.io/badge/openpyxl-028930?style=for-the-badge" alt="OpenPyXL"></a>
</p>

*   **Backend:** Flask (Python)
*   **Web Scraping:** Selenium
*   **Data Analysis:** pandas
*   **Excel Interaction:** openpyxl
*   **ChromeDriver Management:** chromedriver\_autoinstaller

## License

MIT License