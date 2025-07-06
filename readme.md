# <p align="center">Excel and Twitter Data Integration API</p>

<p align="center">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium&logoColor=white" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"></a>
</p>

## Introduction

This project provides an API to integrate Excel automation with Twitter (X) data analysis. It allows users to process queries against Excel spreadsheets using an AI agent and scrape/analyze Twitter data for specific insights. This API targets developers seeking to build data-driven applications that combine social media intelligence with spreadsheet functionality.

## Table of Contents

1.  [Key Features](#key-features)
2.  [Installation Guide](#installation-guide)
3.  [Usage](#usage)
4.  [Environment Variables](#environment-variables)
5.  [Project Structure](#project-structure)
6.  [Technologies Used](#technologies-used)
7.  [License](#license)

## Key Features

*   **Excel Automation via API:**  Execute commands and process queries against Excel spreadsheets using a Flask-based API.
*   **AI-Powered Agent:** Integrates an `excel_agent` (not defined in these snippets) to process queries and automate actions within Excel.
*   **Twitter (X) Data Scraping:** Scrape tweets based on specified search queries using Selenium.
*   **Data Analysis:** Process scraped tweet data, identify emergency-related tweets, and extract location information using Pandas.
*   **Health Check Endpoint:**  Provides an API endpoint to check the status of the Excel connection and API configuration.
*   **Robust Error Handling:** Implements comprehensive error handling to ensure application stability.

## Installation Guide

1.  **Clone the repository:**

    ```bash
    git clone <repository_url>
    cd <repository_name>
    ```

2.  **Install dependencies:**

    ```bash
    pip install flask selenium pandas chromedriver_autoinstaller
    ```

3.  **Set up environment variables:**

    Create a `.env` file in the project root directory and define the necessary environment variables (see [Environment Variables](#environment-variables) section).  Example:

    ```
    EXCEL_CONNECTION_STRING=your_excel_connection_string
    GEMINI_API_KEY=your_gemini_api_key
    CHROME_PROFILE_PATH=/path/to/your/chrome/profile
    SEARCH_QUERY=your_twitter_search_query
    ```

4.  **Run the Flask server:**

    ```bash
    python app.py
    ```

## Usage

The API provides several endpoints for interacting with Excel and triggering data scraping:

*   **`/` (GET):** Renders the main application page.
*   **`/process_query` (POST):** Processes a query against Excel using the AI agent. Requires `query`, `context`, and optionally `image_data` in the request body.
*   **`/autonomous_action` (POST):** Processes a query and executes actions in Excel. Requires `query`, `context`, and optionally `image_data` in the request body.
*   **`/connect_to_excel` (GET):** Attempts to connect to the Excel instance.
*   **`/execute_command` (POST):** Executes a specific command in Excel. Requires `command` in the request body.
*   **`/health_check` (GET):** Checks the status of the API and the Excel connection.

To scrape Twitter data, the `main.py` script can be executed directly:

```bash
python main.py
```

This will initialize a Chrome WebDriver, scrape tweets based on the `SEARCH_QUERY` environment variable, and process the data.

## Environment Variables

The following environment variables are required:

*   `EXCEL_CONNECTION_STRING`: Connection string for the Excel instance.  This allows the API to connect to the desired Excel workbook.
*   `GEMINI_API_KEY`: API key for the Gemini AI model. This is used by the `excel_agent` to process queries.
*   `CHROME_PROFILE_PATH`: Path to the Chrome user profile. This allows Selenium to use a specific Chrome profile for web scraping.
*   `SEARCH_QUERY`: The search query used for scraping tweets from Twitter (X).

## Project Structure

```
.
├── app.py      # Flask application defining API endpoints
├── main.py     # Script for scraping and processing Twitter data
└── README.md   # This file
```

## Technologies Used

<p align="center">
  <a href="#"><img src="https://img.shields.io/badge/Flask-000000?style=for-the-badge&logo=flask&logoColor=white" alt="Flask"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium&logoColor=white" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
</p>

*   **Backend:** Flask (Python)
*   **Web Scraping:** Selenium, `chromedriver_autoinstaller`
*   **Data Analysis:** Pandas
*   **AI Agent:** Gemini (via `excel_agent` - not defined in provided files)

## License

MIT License
