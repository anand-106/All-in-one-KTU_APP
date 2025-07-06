# <p align="center">Excel Data Automation and Twitter Scraper API</p>

<p align="center">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium&logoColor=white" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"></a>
</p>

## Introduction

This project provides a FastAPI-based API for automating Excel tasks using natural language queries and scraping Twitter data to provide context. It allows users to interact with Excel spreadsheets through API calls and retrieves relevant tweets based on specified search criteria, enhancing data analysis and decision-making processes. This project is designed for data analysts, automation engineers, and developers looking to integrate Excel automation and real-time data retrieval into their workflows.

## Table of Contents

1.  [Key Features](#key-features)
2.  [Installation Guide](#installation-guide)
3.  [Usage](#usage)
4.  [API Reference](#api-reference)
5.  [Environment Variables](#environment-variables)
6.  [Project Structure](#project-structure)
7.  [Technologies Used](#technologies-used)
8.  [License](#license)

## Key Features

*   **Excel Automation via Natural Language:** Enables users to perform complex Excel operations using natural language queries through the `excel_agent` (external component).
*   **Twitter Data Scraping:** Scrapes real-time tweets based on a defined `SEARCH_QUERY` using Selenium.
*   **Emergency Tweet Filtering:** Filters scraped tweets based on predefined emergency keywords and location mentions to identify relevant information.
*   **REST API Endpoints:** Provides a set of API endpoints for processing user queries, executing commands, and managing Excel connections.
*   **Health Check Endpoint:** Monitors the status of the API, Excel connection, and Gemini configuration.
*   **Autonomous Excel Actions:** Executes actions in Excel based on user queries and context.

## Installation Guide

1.  **Clone the repository:**

    ```bash
    git clone <repository_url>
    cd <repository_name>
    ```

2.  **Install dependencies:**

    ```bash
    pip install -r requirements.txt
    ```
    Make sure you have Python 3.7+ installed.

3.  **Set up environment variables:**

    Create a `.env` file in the project root and add the following variables:
    ```
    # Example (replace with actual values)
    EXCEL_AGENT_API_KEY=your_excel_agent_api_key
    SEARCH_QUERY=emergency+alerts
    CHROME_PROFILE_PATH=/path/to/chrome/profile
    CHROME_PROFILE_NAME=YourProfileName
    ```

    **Note:** The `excel_agent` requires specific configuration details, which are assumed to be managed externally. Update the `.env` with the proper information.

4.  **Run the FastAPI server:**

    ```bash
    python main.py
    python app.py # you might need to run this in another terminal
    ```
    Access the API at `http://localhost:8000` (or the port specified in your configuration).

## Usage

The API provides several endpoints for interacting with Excel and retrieving data.

*   **Processing User Queries:** Send a POST request to `/process_query` with a JSON payload containing the user's query.
*   **Executing Commands:** Send a POST request to `/execute_command` with a JSON payload containing the command to execute in Excel.
*   **Autonomous Actions:** Send a POST request to `/autonomous_action` with a JSON payload containing the query and context.
*   **Connecting to Excel:** Send a GET request to `/connect_to_excel` to establish a connection to an Excel instance.
*   **Health Check:** Send a GET request to `/health_check` to check the status of the API and its dependencies.

The `scrape_tweets` function in `main.py` is executed separately to gather Twitter data, which can then be used as context for the Excel agent.

## API Reference

(This section would contain detailed documentation of the API endpoints, request/response formats, etc. if more details were available from the code summaries.)

**Example Endpoints:**

*   **/process\_query**
    *   **Method:** POST
    *   **Request Body:** `{"query": "Summarize sales data"}`
    *   **Response Body:** `{"response": "Sales data summary"}`
*   **/health\_check**
    *   **Method:** GET
    *   **Response Body:** `{"api_status": "OK", "excel_connection": "Connected"}`

## Environment Variables

*   `EXCEL_AGENT_API_KEY`: API key for authenticating with the `excel_agent`.  This is a placeholder, adjust based on the actual `excel_agent` API requirements.
*   `SEARCH_QUERY`: The search query used to scrape tweets from Twitter (e.g., "emergency+alerts").
*   `CHROME_PROFILE_PATH`: Path to the Chrome user profile directory.
*   `CHROME_PROFILE_NAME`: Name of the Chrome profile to use.
*   `EMERGENCY_KEYWORDS`: Comma separated list of keywords to detect emergencies in tweets
*   `LOCATIONS`: Comma separated list of locations to search for in tweets

## Project Structure

```
.
├── app.py       # Flask API endpoints
├── main.py      # Twitter scraping and data processing functions
├── README.md    # Project documentation
└── requirements.txt # Project dependencies
```

## Technologies Used

<p align="left">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium&logoColor=white" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"></a>
  <a href="#"><img src="https://img.shields.io/badge/chromedriver-0079D6?style=for-the-badge&logo=google-chrome&logoColor=white" alt="ChromeDriver"></a>
</p>

*   **Backend:** FastAPI (Python)
*   **Web Scraping:** Selenium, `chromedriver_autoinstaller`
*   **Data Processing:** Pandas
*   **Other:** Flask, Regular Expressions (re)

## License

MIT License