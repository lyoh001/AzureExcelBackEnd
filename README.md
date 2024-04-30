# Microsoft Graph API Async Downloader

## Features
- Asynchronously downloads files from Microsoft Graph API.
- Utilizes FastAPI for building the web application.
- Implements middleware for CORS handling.
- Uses asyncio and aiohttp for efficient asynchronous operations.

## Prerequisites
Before running the application, ensure you have the following installed:
- Python 3.8 or higher
- [pip](https://pip.pypa.io/en/stable/installing/) (Python package installer)
- [uvicorn](https://www.uvicorn.org/) (ASGI server)
- [dotenv](https://pypi.org/project/python-dotenv/) (For loading environment variables)
- [openpyxl](https://pypi.org/project/openpyxl/) (For working with Excel files)
- [pandas](https://pandas.pydata.org/) (For data manipulation and analysis)
- [aiohttp](https://docs.aiohttp.org/en/stable/) (For making asynchronous HTTP requests)
- [fastapi](https://fastapi.tiangolo.com/) (For building APIs with Python)

## Installation
1. Clone this repository to your local machine using `git clone`.
2. Navigate to the project directory.
3. Install the dependencies using the following command:

```
pip install -r requirements.txt

```

## Usage
1. Ensure you have set up your environment variables by creating a `.env` file in the project root directory. Refer to `.env.example` for required variables.
2. Run the application using the following command:
```
python app.py

```
3. Access the API endpoints using a web browser or an API client like [Postman](https://www.postman.com/).

## Endpoints
- `GET /status`: Returns the status of the application.
- `GET /load`: Initiates the file download process from Microsoft Graph API.
- `POST /update`: Updates the application (not implemented).

## Contributing
Contributions to this project are welcome. To contribute, follow these steps:
1. Fork the repository.
2. Create a new branch (`git checkout -b feature/fooBar`).
3. Make your changes and commit them (`git commit -am 'Add some fooBar'`).
4. Push to the branch (`git push origin feature/fooBar`).
5. Create a new Pull Request.

## License
This project is licensed under the MIT License.
