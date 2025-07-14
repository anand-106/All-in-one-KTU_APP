# <p align="center">TweetScraper-Excel Integration</p>

<p align="center">
    <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
    <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
    <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium&logoColor=white" alt="Selenium"></a>
    <a href="#"><img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"></a>
</p>

## Introduction

This project automates data extraction from Twitter/X, specifically scraping tweets based on a defined search query. It then processes this data and allows interaction with an Excel application for further analysis or command execution. The system targets users needing automated social media data collection and integration with existing Excel workflows.

## Table of Contents

1.  [Key Features](#key-features)
2.  [Installation Guide](#installation-guide)
3.  [Usage](#usage)
4.  [Environment Variables](#environment-variables)
5.  [Project Structure](#project-structure)
6.  [Technologies Used](#technologies-used)
7.  [License](#license)

## Key Features

*   **Tweet Scraping:** Scrapes tweets from Twitter/X based on a user-defined search query using Selenium.
*   **Data Processing:** Processes scraped tweet data using pandas, identifying emergency-related tweets and locations.
*   **Excel Integration:** Connects to an Excel instance and executes commands based on user queries or autonomous actions.
*   **REST API:** Provides a Flask-based REST API for interacting with the scraping and Excel functionalities.
*   **Error Handling:** Robust error handling to prevent application crashes and provide informative error messages.

## Installation Guide

1.  **Clone the repository:**

    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2.  **Install the dependencies:**

    ```bash
    pip install -r requirements.txt # Add a requirements.txt to the repo, if one is needed.  It is good practice.
    ```

3.  **Set up environment variables:**

    Create a `.env` file in the project root directory and add the following variables:

    ```
    # Example .env file
    CHROME_PROFILE_PATH=/path/to/chrome/profile
    CHROME_PROFILE_NAME=YourProfileName
    SEARCH_QUERY=YourSearchQuery
    ```
    *(Note: the code chunk summary implies these variables; create a `requirements.txt` with necessary packages if one doesn't exist)*

4.  **Run the Flask API:**

    ```bash
    python app.py
    ```

## Usage

The Flask API provides several endpoints:

*   `/`: Renders the `index.html` (if applicable).
*   `/process_query`: Processes a user query against the Excel application.  Send a JSON payload with a `query` field.
*   `/autonomous_action`: Processes a query with context data, triggering autonomous actions in Excel. Send a JSON payload.
*   `/execute_command`: Executes a specific command in the Excel application. Send a JSON payload with a `command` field.
*   `/health_check`: Checks the health of the API and the Excel connection.
*   `/connect_to_excel`: Connects to an Excel instance.

Example using `curl`:

```bash
curl -X POST -H "Content-Type: application/json" -d '{"query": "your query"}' http://localhost:5000/process_query
```

## Environment Variables

The following environment variables are required:

*   `CHROME_PROFILE_PATH`: The path to the Chrome profile directory.
*   `CHROME_PROFILE_NAME`: The name of the Chrome profile.
*   `SEARCH_QUERY`: The search query used for scraping tweets from Twitter/X.

*It is good practice to store the `SEARCH_QUERY` in the environment variables for flexibility*

## Project Structure

```
.
├── app.py       # Flask API application
├── main.py      # Tweet scraping and data processing logic
└── readme.md    # This README file
```

## Technologies Used

<p align="center">
    <a href="#"><img src="https://img.shields.io/badge/Flask-000000?style=for-the-badge&logo=flask&logoColor=white" alt="Flask"></a>
    <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium&logoColor=white" alt="Selenium"></a>
    <a href="#"><img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"></a>
    <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
</p>

*   **Backend:** Flask (Python)
*   **Web Scraping:** Selenium
*   **Data Processing:** pandas
*   **Automation:** webdriver\_manager

## License

MIT License