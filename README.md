# MLSA Certificate Generator

This project is a Flask application that allows users to generate personalized certificates for participants of an event for Microosft Learn Student Ambassadors. Certificates are generated based on a provided template and participant details from a CSV file. The generated certificates are then converted to PDF format and zipped for download.


## Live Url:

http://mlsa-certificate.onrender.com

## Features

- Upload a CSV file with participant names.
- Specify the event name and ambassador name.
- Generate certificates in PDF format.
- Download all certificates in a single ZIP file.

## Installation

To get started with this project, follow these steps:

1. **Clone the repository:**
    ```sh
    git clone https://github.com/Ramseyxlil/MLSA_certificate.git
    cd MLSA_certificate
    ```

2. **Create and activate a virtual environment:**
    ```sh
    python3 -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

3. **Install the required dependencies:**
    ```sh
    pip install -r requirements.txt
    ```

4. **Set up ConvertAPI:**
    - Sign up at [ConvertAPI](https://www.convertapi.com/) to get your API secret key.
    - Replace `'yQBA8IqyogHySlgp'` in the script with your actual API secret key.

## Usage

1. **Run the Flask application:**
    ```sh
    python app.py
    ```

2. **Open your browser and go to `http://127.0.0.1:5000/` to access the application.**

3. **Upload a CSV file:**
    - The CSV file should contain a column named `Name` with the names of participants.

4. **Fill in the event name and ambassador name.**

5. **Generate and download the certificates:**
    - After uploading the file and submitting the form, a ZIP file containing all the certificates in PDF format will be available for download.

## Project Structure

- `app.py`: Main Flask application script.
- `templates/`: Folder containing the HTML template for file upload (`upload.html`).
- `uploads/`: Folder for storing uploaded CSV files (created automatically).
- `certificates/`: Folder for storing generated certificates (created automatically).
- `zips/`: Folder for storing the ZIP file of certificates (created automatically).
- `certificate_template.docx`: The template file for the certificates. 

## Code Overview

### Flask Routes

- **`/`**: Renders the file upload form.
- **`/upload`**: Handles the file upload and certificate generation.
- **`/download/<path:path>`**: Serves the ZIP file for download and deletes it after the download is complete.

### Functions

- **`apply_font_style(run, font_size, color, bold)`**: Applies the specified font style to a text run in the document.
- **`generate_certificate(participant_name, event_name, ambassador_name)`**: Generates a certificate for a single participant.
- **`generate_certificates(file_path, event_name, ambassador_name)`**: Processes the CSV file and generates certificates for all participants.
- **`convert_to_pdf(docx_path)`**: Converts a DOCX file to PDF using ConvertAPI.
- **`create_zip(event_name)`**: Creates a ZIP file containing all the generated PDF certificates.

## Example CSV File

```csv
Name
John Doe
Jane Smith
License
```

 This project is licensed under the MIT License. See the LICENSE file for details.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

Acknowledgements

ConvertAPI for the document conversion API.
The Flask framework for making web development in Python easy and fun.
