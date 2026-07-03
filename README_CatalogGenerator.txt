Catalog Generator
=================

What this app does
------------------
Creates a catalog PDF from an Excel file.
It can also add a watermark to an already generated catalog PDF without needing the Excel file again.

You choose:
- Excel file
- Product name column
- Image URL column
- Optional header image
- Product grid columns and rows, such as 3 columns x 5 rows or 1 column x 2 rows
- Optional watermark image for each product
- Downloaded image folder
- PDF save folder
- PDF file name
- Existing PDF for making a watermarked copy

Layout
------
- A4 page
- Flexible product grid size using separate columns and rows fields
- No line under the header
- No card borders
- No side or bottom margin
- Original downloaded/source images are kept locally
- The PDF embeds optimized high-quality copies to keep the PDF file size small

Files
-----
catalog_generator_app.py
    The reusable desktop app source.

build_exe.bat
    Builds CatalogGenerator.exe using PyInstaller.

How to make the EXE
-------------------
1. Put catalog_generator_app.py and build_exe.bat in the same folder.
2. Double-click build_exe.bat.
3. Wait for it to finish.
4. Use the generated CatalogGenerator.exe.

Recommended fast workflow
-------------------------
1. Choose the Excel file.
2. Enter/select the product name column.
   You may use the header name or an Excel letter like B.
3. Enter/select the image URL column.
   You may use the header name or an Excel letter like T.
4. Choose a downloaded image folder.
5. Enter the number of product columns and rows.
6. Optional: choose a watermark image and turn on watermark.
7. Choose the PDF save folder and file name.
8. Click Download Images First.
9. Click Generate PDF.

The app saves each downloaded image locally using the product name plus a URL fingerprint.
On the next PDF generation, it reuses those files instead of downloading the same URLs again.
During PDF generation, it creates a _pdf_optimized folder with compressed copies sized for the PDF.
This keeps the PDF much smaller while preserving good visual quality.
After the PDF is created, a popup lets you open the PDF or open its folder.
The generated PDF folder stays the same across new Excel uploads until you change it manually.
The PDF file name updates from the selected Excel file name to help avoid accidental overwrites.
If any product image is skipped, the Activity log shows the first skipped rows and the app saves
a skipped image report next to the PDF named like YourCatalog_skipped_images.csv.

Add watermark to an existing PDF
--------------------------------
Use this when you already created a plain catalog PDF and want a second watermarked copy.

1. Choose the same product grid used by the original PDF.
2. Choose the watermark image.
3. Choose the generated PDF folder.
4. Select the existing plain PDF.
5. The app auto-detects the total number of products from the existing PDF.
   You can correct the value manually if needed.
   This prevents watermarking blank slots on the last page.
6. Enter the watermarked PDF file name.
7. Click Create Watermarked Copy.

The app creates a small optimized watermark cache beside the watermarked PDF so the output file
does not become unnecessarily large.

Notes
-----
- The first build may take a few minutes.
- Internet is needed while building if PyInstaller is not already installed.
- Internet is also needed while generating a catalog if your Excel file uses online image URLs.
- If you want the exact current header, choose your header_cropped.png file in the app.
