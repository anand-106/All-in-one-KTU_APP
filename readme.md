# <p align="center">Web Scraping, Data Analysis, and Excel Automation API</p>

<p align="center">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas" alt="Pandas"></a>
</p>

## Table of Contents

1.  [Description](#description)
2.  [Key Features](#key-features)
3.  [Installation Guide](#installation-guide)
4.  [Usage](#usage)
5.  [API Reference](#api-reference)
6.  [Environment Variables](#environment-variables)
7.  [Project Structure](#project-structure)
8.  [Technologies Used](#technologies-used)
9.  [License](#license)

## Description

This project combines web scraping, data analysis, and Excel automation into a single API. It leverages Selenium to scrape data from websites (specifically Twitter/X), Pandas for data processing and analysis, and a Flask API to expose these functionalities. The core goal is to scrape emergency-related information, analyze it, and interact with Excel based on user queries. The excel_agent is responsable for maintaining the excel connection.

## Key Features

*   **Web Scraping with Selenium:** Scrapes tweets from Twitter/X based on a predefined search query, extracting text, user information, timestamps, and URLs.
*   **Dynamic Content Handling:**  Handles dynamically loaded content on web pages to ensure all relevant data is scraped.
*   **Data Analysis with Pandas:** Processes raw tweet data, identifies emergency-related tweets based on keywords and location, and creates filtered datasets.
*   **Flask API:** Provides endpoints to process user queries, execute commands in Excel, and perform health checks.
*   **Excel Automation:** Enables interaction with Excel, processing user queries to perform actions and automate tasks. The excel_agent is responsible for keeping the excel connection open.
*   **Error Handling:** Implements robust error handling using `try...except` blocks to prevent application crashes and provide informative error messages.
*   **Health Check Endpoint:** Includes a health check endpoint to verify the API's status and the connection to Excel.

## Installation Guide

1.  **Clone the repository:**

    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2.  **Create a virtual environment (recommended):**

    ```bash
    python3 -m venv venv
    source venv/bin/activate  # On Linux/macOS
    venv\Scripts\activate  # On Windows
    ```

3.  **Install dependencies:**

    ```bash
    pip install -r requirements.txt #if a requirements.txt exists, otherwise:
    pip install selenium pandas flask chromedriver_autoinstaller openpyxl
    ```

4.  **Set up environment variables:**

    Create a `.env` file in the project root directory and define the following environment variables:

    ```
    CHROME_PROFILE_PATH=<path_to_your_chrome_profile>
    CHROME_PROFILE_NAME=<your_chrome_profile_name>
    SEARCH_QUERY=<twitter_search_query>
    EMERGENCY_KEYWORDS=<comma_separated_emergency_keywords>
    LOCATIONS=<comma_separated_locations>
    SCROLL_ITERATIONS=<number_of_scroll_iterations>
    ```

    **Note:** Ensure your Chrome profile path and name are correct. Create a Chrome profile if you don't have one. The `SEARCH_QUERY`, `EMERGENCY_KEYWORDS` and `LOCATIONS` are used for web scraping.

5.  **Run the Flask server:**

    ```bash
    python app.py
    ```

## Usage

The API provides several endpoints for interacting with the web scraping, data analysis, and Excel automation functionalities.

*   **/:** Renders the main page of the application (index.html).
*   **/process_query:** Processes user queries and returns an AI response based on the Excel context.
*   **/autonomous_action:** Processes user queries and autonomously executes actions in Excel.
*   **/execute_command:** Executes a specific command in Excel.
*   **/health_check:** Checks the API's status and the connection to Excel.
*   **/connect:** Connect to excel.
The excel_agent is responsable for managing the excel connection.

To interact with the API, send HTTP requests to the appropriate endpoints with the required data. For example, to process a query, you can send a POST request to `/process_query` with the query, Excel context, and image data in the request body.

## API Reference

### `/process_query`

*   **Method:** POST
*   **Request Body:** JSON object containing the query, Excel context, and image data.
    ```json
    {
      "query": "Your query here",
      "excel_context": "Excel context here",
      "image_data": "Base64 encoded image data (optional)"
    }
    ```
*   **Response:** JSON object containing the AI response or an error message.

### `/autonomous_action`

*   **Method:** POST
*   **Request Body:** JSON object containing the query, Excel context, and image data.
    ```json
    {
      "query": "Your query here",
      "excel_context": "Excel context here",
      "image_data": "Base64 encoded image data (optional)"
    }
    ```
*   **Response:** JSON object containing the result of the autonomous action or an error message.

### `/execute_command`

*   **Method:** POST
*   **Request Body:** JSON object containing the command to execute.
    ```json
    {
      "command": "Excel command here"
    }
    ```
*   **Response:** JSON object containing the result of the command execution or an error message.

### `/health_check`

*   **Method:** GET
*   **Response:** JSON object containing the API status, Excel connection status, and workbook name (if connected).
    ```json
    {
      "status": "OK",
      "excel_connected": true,
      "workbook_name": "Workbook1.xlsx"
    }
    ```

## Environment Variables

| Variable              | Description                                                                    |
| --------------------- | ------------------------------------------------------------------------------ |
| `CHROME_PROFILE_PATH` | Path to your Chrome profile.                                                   |
| `CHROME_PROFILE_NAME` | Name of your Chrome profile.                                                   |
| `SEARCH_QUERY`        | The search query used for scraping tweets from Twitter/X.                      |
| `EMERGENCY_KEYWORDS`  | Comma-separated list of keywords used to identify emergency-related tweets.   |
| `LOCATIONS`           | Comma-separated list of locations to filter tweets by.                        |
| `SCROLL_ITERATIONS`   | The number of times to scroll the page while scraping data.                   |

## Project Structure

```
.
├── app.py        # Flask API application
├── main.py       # Web scraping functions
└── readme.md     # This file
```

## Technologies Used

<p align="center">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas" alt="Pandas"></a>
  <a href="#"><img src="https://img.shields.io/badge/openpyxl-057D57?style=for-the-badge" alt="openpyxl"></a>
</p>

*   **Backend:** Flask (Python)
*   **Web Scraping:** Selenium (Python)
*   **Data Analysis:** Pandas (Python)
*   **Excel Interaction:** openpyxl (Python)
*   **ChromeDriver Management:** chromedriver\_autoinstaller (Python)

## License

MIT License

<p align="center">
  <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="MIT License"></a>
</p>