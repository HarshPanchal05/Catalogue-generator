# Catalogue Generator

A Python-based catalogue generator that creates product catalogue PDFs using product names and image URLs from an Excel file.

## Features

* Generate product catalogue pages automatically
* Add product name and product image
* Supports multiple products per page
* Saves time compared to manual drag-and-drop design
* Useful for product catalogues, spare parts catalogues, and business listings

## Tech Stack

* Python
* Pandas
* ReportLab / PDF generation libraries
* Excel input file

## Input Format

Prepare an Excel file with the following columns:

| Product Name | Image URL                      |
| ------------ | ------------------------------ |
| Product 1    | https://example.com/image1.jpg |
| Product 2    | https://example.com/image2.jpg |

## Installation

```bash
git clone https://github.com/HarshPanchal05/Catalogue-generator.git
cd Catalogue-generator
pip install -r requirements.txt
```

## Usage

```bash
python catalogue_generator.py
```

After running the script, the generated catalogue PDF will be saved in the output folder.

## Project Purpose

This project was created to automate the process of making product catalogues. Instead of manually placing each product image and name, the script reads data from an Excel file and generates a clean catalogue layout automatically.

## Author

Harsh Panchal

GitHub: [HarshPanchal05](https://github.com/HarshPanchal05)
