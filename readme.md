# <p align="center">Automated Twitter and Excel Integration</p>

<p align="center">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas" alt="Pandas"></a>
</p>

## Introduction

This project automates interactions between Twitter and Excel, enabling users to perform actions based on Twitter data and user queries. It leverages web scraping to extract tweets, processes them, and interacts with an Excel backend via a Flask API. The target users are data analysts, researchers, and anyone who needs to integrate real-time Twitter data with Excel-based workflows.

## Table of Contents

1.  [Key Features](#key-features)
2.  [Installation Guide](#installation-guide)
3.  [Usage](#usage)
4.  [Environment Variables](#environment-variables)
5.  [Project Structure](#project-structure)
6.  [Technologies Used](#technologies-used)
7.  [License](#license)

## Key Features

*   **Twitter Scraping:** Extracts tweets based on user-defined search queries using Selenium. Utilizes Chrome profiles for realistic scraping behavior.
*   **Data Processing:** Identifies and filters tweets related to emergencies and specific locations using regular expressions and Pandas.
*   **Excel Integration:** Executes commands and processes queries within Excel via an `excel_agent`.
*   **API Endpoints:** Provides a Flask API for interacting with the system, including processing queries, executing commands, and checking system health.
*   **Autonomous Actions:** Can execute autonomous actions in Excel based on user queries and context.

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

    *(Note: A `requirements.txt` file was not in the code chunk summaries, so create it with the following dependencies: `flask`, `selenium`, `pandas`, `webdriver_manager`)*

3.  **Set up environment variables:**

    Create a `.env` file in the project root and set the following variables:

    ```
    EXCEL_AGENT_CONFIG=<path_to_excel_agent_config> # Path to excel agent configuration file
    #Example:
    # AUTH0_DOMAIN=your_auth0_domain
    # AUTH0_CLIENT_ID=your_auth0_client_id
    # AUTH0_CLIENT_SECRET=your_auth0_client_secret
    DATABASE_URI=your_database_uri
    ```

    **Important:** Replace placeholder values with your actual credentials and configuration paths. The `EXCEL_AGENT_CONFIG` variable points to the config the excel agent requires to function.

4.  **Run the FastAPI server:**

    ```bash
    python app.py
    ```

## Usage

The application exposes several API endpoints for interacting with the system. You can send requests to these endpoints to process queries, execute commands in Excel, and retrieve data.

*   **/process_query:** Accepts a user query and returns the AI's response after processing it through the `excel_agent`.
*   **/autonomous_action:** Accepts a query and context data. Executes an autonomous action in Excel and returns the result.
*   **/execute_command:** Executes a specific command in Excel.
*   **/health_check:** Checks the health status of the API and its dependencies.
*   **/connect_to_excel:** Attempts to connect to the Excel instance.

Example usage (using `curl`):

```bash
curl -X POST -H "Content-Type: application/json" -d '{"query": "Summarize sales data"}' http://localhost:5000/process_query
```

## Environment Variables

The following environment variables are required for the application to function correctly:

*   `EXCEL_AGENT_CONFIG`:  Path to the Excel agent configuration file.  This file defines how the application interacts with Excel.
*   `DATABASE_URI`:  The URI for the database connection (if applicable for storing data). (Not used in current code but included for potential future use based on usual project structures.)

## Project Structure

```
.
├── app.py      # Flask API endpoints and logic
├── main.py     # Twitter scraping and data processing functions
└── readme.md   # Project documentation
```

## Technologies Used

<p align="center">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas" alt="Pandas"></a>
</p>

*   **Backend:** Flask (Python)
*   **Web Scraping:** Selenium, WebDriver Manager
*   **Data Processing:** Pandas, Regular Expressions
*   **API:** Flask
*   **Excel Integration:** Custom `excel_agent` (implementation not provided)

## License

MIT License