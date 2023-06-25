import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
import pandas as pd
import os

def extract_images_from_excel(file_path, output_folder):
    # Load the workbook
    wb = openpyxl.load_workbook(file_path)
    
    # List to store image data
    image_data = []

    # Iterate through all worksheets
    for ws in wb.worksheets:
        # Iterate through drawings in the worksheet
        for drawing in ws._drawing:
            # Check if drawing is an instance of openpyxl Image
            if isinstance(drawing, Image):
                # Save image to output folder
                image_path = os.path.join(output_folder, drawing.image.filename)
                with open(image_path, 'wb') as f:
                    f.write(drawing.image.fp.read())
                
                # Store image data
                image_data.append({
                    'Excel File Name': os.path.basename(file_path),
                    'Sheet Name': ws.title,
                    'Image Name': drawing.image.filename,
                    'Row': drawing.coordinates[1],
                    'Column': drawing.coordinates[0]
                })
                
            # Check if drawing is an instance of openpyxl SpreadsheetDrawing
            elif isinstance(drawing, SpreadsheetDrawing):
                # Save image to output folder
                image_path = os.path.join(output_folder, drawing.imagePart.img._rels[drawing.imagePart._relId].relId + '.png')
                with open(image_path, 'wb') as f:
                    f.write(drawing.imagePart.image.imageData)
                
                # Store image data
                image_data.append({
                    'Excel File Name': os.path.basename(file_path),
                    'Sheet Name': ws.title,
                    'Image Name': drawing.imagePart.img._rels[drawing.imagePart._relId].relId + '.png',
                    'Row': drawing.coordinates[1],
                    'Column': drawing.coordinates[0]
                })

    # Create DataFrame
    df = pd.DataFrame(image_data)
    return df

# Use the function
df = extract_images_from_excel('path_to_your_file.xlsx', 'output_folder')
print(df)
