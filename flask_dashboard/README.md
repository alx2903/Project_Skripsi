# Flask Dashboard App

This guide provides step-by-step instructions on how to set up and run the Flask-based dashboard application.

## Prerequisites

Before running the application, ensure you have the following installed:

- Python 3.8 or later ([Download Python](https://www.python.org/downloads/))
- pip (Python package manager)
- Virtual environment (recommended for dependency management)

## Installation Steps

### 1. Clone the Repository

If you received the project in a ZIP file, extract it. If you have access to a repository, clone it using:
```sh
git clone <repository_url>
cd flask_dashboard_app  # Change to the project directory
```

### 2. Set Up a Virtual Environment (Recommended)

Creating a virtual environment helps manage dependencies:
```sh
python -m venv venv  # Create a virtual environment
```

Activate the virtual environment:
- **Windows:**
  ```sh
  venv\Scripts\activate
  ```
- **Mac/Linux:**
  ```sh
  source venv/bin/activate
  ```

### 3. Install Dependencies

Run the following command to install the required packages:
```sh
pip install -r requirements.txt
```

### 4. Run the Application

Start the Flask app using:
```sh
python app.py
```

The application will start running on `http://127.0.0.1:5000/`.

### 5. Upload and Analyze Data

1. Open a web browser and visit `http://127.0.0.1:5000/`.
2. Upload an Excel file (`.xlsx`) containing the required data.
3. The dashboard will display various analytical insights.

### 6. (Optional) Run Flask in Production Mode

If you want to deploy this app on a server, consider using Gunicorn (for Linux/macOS) or Waitress (for Windows):

- **Linux/macOS:**
  ```sh
  pip install gunicorn
  gunicorn -w 4 -b 0.0.0.0:8000 app:app
  ```

- **Windows:**
  ```sh
  pip install waitress
  waitress-serve --listen=0.0.0.0:8000 app:app
  ```

### 7. Deactivating Virtual Environment

After running the app, deactivate the virtual environment:
```sh
deactivate
```

## Troubleshooting

- **Python Not Found:** Ensure Python is installed and added to the system PATH.
- **Missing Dependencies:** Run `pip install -r requirements.txt`.
- **Port Already in Use:** Change the port using `python app.py --port=5001`.

## Folder Structure
```
flask_dashboard_app/
|-- app.py                 # Main Flask application
|-- templates/
|   |-- index.html         # File upload page
|   |-- dashboard.html     # Dashboard visualization
|   |-- loading.html       # Loading page
|   |-- confirm.html       # Confirm page
|-- static/
|   |-- css/               # CSS files (if any)
|   |-- js/                # JavaScript files (if any)
|-- uploads/               # Folder where uploaded files are stored
|-- requirements.txt       # Python dependencies
|-- README.md              # This guide
```

## Conclusion
This guide provides all necessary steps to run the Flask dashboard application. If you encounter issues, refer to the troubleshooting section or contact the project maintainer.

