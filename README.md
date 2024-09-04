![Certificates Preview](https://raw.githubusercontent.com/Anuswar/certificate-generator/main/preview.jpg)

# Certificate Generator 🏆

Welcome to the Certificate Generator repository! This project automates the creation of personalized certificates in PDF format using a PowerPoint template. With a single script, you can quickly generate certificates for multiple recipients, ensuring consistent design and format.

### ✨ Key Features of the Script:
- Automatically replaces placeholders with recipient names while preserving the text style and formatting.
- Generates certificates for all names listed in a text file.
- Saves the certificates as PDFs, ready for distribution.
- Ensures the output directory exists, creating it if necessary.

## ⚙️ Installation

1. **Clone the repository to your local machine:**
    ```sh
    git clone https://github.com/Anuswar/certificate-generator.git
    cd certificate-generator
    ```

2. **Set up the environment:**
    - Ensure you have Python installed.
    - Install the necessary Python packages:
      ```sh
      pip install python-pptx comtypes
      ```

3. **Prepare your files:**
    - Place your `certificate_template.pptx` PowerPoint template in the repository directory.
    - Create a `names.txt` file with each recipient's name on a new line.

4. **Run the script:**
    ```sh
    python certificate-generator.py
    ```

## 📂 File Structure

The repository includes the following files:

```
certificate-generator/
├── generate_certificates.py  # Script to generate certificates
├── certificate_template.pptx # PowerPoint template for the certificates
├── names.txt                 # List of recipient names
├── README.md                 # This README file
├── LICENSE.md                # License for the repository
```

## 🛠️ Tech Stack

This repository utilizes the following technologies and tools:

- **Python**: The main scripting language used for generating certificates.
- **python-pptx**: A Python library for creating and updating PowerPoint files.
- **comtypes**: A Python library to interface with COM objects, used here for converting PPTX files to PDFs.
- **PowerPoint**: Used as the template base for certificates.
- **PDF**: The final output format for the certificates, ensuring easy distribution and printing.

## 🤝 Contributing

Contributions are welcome! If you find any issues, have suggestions, or want to add features, please follow these steps:

1. **Fork the repository.**
2. **Create a new branch** for your feature or bug fix.
3. **Make your changes and commit them** with descriptive messages.
4. **Push your changes** to your fork.
5. **Open a pull request** to the `main` branch of the original repository.

## 📄 License

This project is licensed under the [MIT License](LICENSE.md), allowing you to use, modify, and distribute the code freely.
