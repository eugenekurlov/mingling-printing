import io
import math
import os

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.lib.utils import ImageReader
import pymupdf


def add_images_to_pdf_in_grid(
        output_path=None,
        image_paths=None,
        columns=1,
        rows=1,
        angles=None,
        orientation="portrait",
        page_margin=0,
        image_margin=0):
    """
    Creates a PDF file with images arranged in a grid format on each page, with configurable rotation,
    margins, and orientation.

    Parameters:
    - output_path (str or None): The file path where the PDF should be saved. If None, the PDF is saved to a buffer.
    - image_paths (list of str): A list of file paths to the images to include in the PDF.
    - columns (int): Number of columns in the grid layout.
    - rows (int): Number of rows in the grid layout.
    - angles (list of int or None): A list of rotation angles (in degrees) for each image.
      If None, no rotation is applied. If the list has fewer items than `image_paths`, the remaining images will not be rotated.
    - orientation (str): Page orientation, either "portrait" or "landscape".
    - page_margin (int or float): The margin in points between the content and the page edges.
    - image_margin (int or float): The margin in points between each image within the grid.

    Returns:
    - BytesIO or None: If `output_path` is None, returns an in-memory BytesIO object containing the PDF.
      If `output_path` is specified, saves the PDF to the file and returns None.
    """
    if orientation == "landscape":
        page_width, page_height = landscape(A4)
    else:
        page_width, page_height = portrait(A4)

    # Set up the canvas to write to the provided output (file or buffer)
    if output_path:
        # Set up the canvas and page size based on orientation
        c = canvas.Canvas(output_path, pagesize=(page_width, page_height))
    else:
        output_buffer = io.BytesIO()
        c = canvas.Canvas(output_buffer, pagesize=(page_width, page_height))

    if image_paths is None:
        raise ValueError("image_paths must be provided and cannot be empty.")

    # Calculate usable dimensions with page margin
    usable_width = page_width - 2 * page_margin
    usable_height = page_height - 2 * page_margin

    # Calculate cell dimensions within the usable area
    cell_width = usable_width / columns
    cell_height = usable_height / rows

    # Set default angle list if not provided
    if angles is None:
        angles = [0] * len(image_paths)
    elif len(angles) < len(image_paths):
        angles += [0] * (len(image_paths) - len(angles))  # Fill extra angles with 0 if fewer angles than images

    # Iterate over images, placing each in the grid
    for i, image_path in enumerate(image_paths):
        img = ImageReader(image_path)
        img_width, img_height = img.getSize()

        # Get rotation angle for current image
        angle = angles[i]
        angle_rad = math.radians(angle)

        # Calculate rotated bounding box
        rotated_width = abs(img_width * math.cos(angle_rad)) + abs(img_height * math.sin(angle_rad))
        rotated_height = abs(img_width * math.sin(angle_rad)) + abs(img_height * math.cos(angle_rad))

        # Calculate effective cell dimensions with image margin
        effective_cell_width = cell_width - 2 * image_margin
        effective_cell_height = cell_height - 2 * image_margin

        # Determine scaling factor to fit image within effective cell dimensions
        scale_factor = min(effective_cell_width / rotated_width, effective_cell_height / rotated_height, 1)
        final_width = img_width * scale_factor
        final_height = img_height * scale_factor

        # Determine grid position
        col = i % columns
        row = (i // columns) % rows

        # Create new page if needed
        if i > 0 and i % (columns * rows) == 0:
            c.showPage()

        # Position the image within the page, accounting for margins
        x_pos = page_margin + col * cell_width + image_margin + (effective_cell_width - final_width) / 2
        y_pos = page_height - (page_margin + (row + 1) * cell_height) + image_margin + (
                    effective_cell_height - final_height) / 2

        # Draw the image with rotation and scaling
        c.saveState()
        c.translate(x_pos + final_width / 2, y_pos + final_height / 2)
        c.rotate(angle)
        c.drawImage(image_path, -final_width / 2, -final_height / 2, width=final_width, height=final_height,
                    preserveAspectRatio=True)
        c.restoreState()

    # Finalize PDF
    c.showPage()
    c.save()

    if output_path:
        return output_path
    else:
        # If using an output buffer, return its contents to the start for reading
        output_buffer.seek(0)
        return output_buffer


def create_pdf_with_best_orientation_images(
        output_path=None,
        image_paths=None,
        columns=1,
        rows=1,
        orientation="auto",
        page_margin=0,
        image_margin=0):
    """
    Creates a PDF with images arranged in a grid on each page, automatically determining
    the best page orientation (portrait or landscape) based on image aspect ratios.

    Parameters:
    - output_path (str or None): The file path where the PDF should be saved. If None, the PDF is saved to an in-memory buffer.
    - image_paths (list of str): A list of file paths to the images to include in the PDF.
    - columns (int): Number of columns in the grid layout.
    - rows (int): Number of rows in the grid layout.
    - orientation (str): Page orientation, either "portrait", "landscape", or "auto". If "auto", the function
      automatically selects portrait or landscape based on the layout and image aspect ratios.
    - page_margin (int or float): Margin in points between the content and the page edges.
    - image_margin (int or float): Margin in points between each image within the grid.

    Returns:
    - BytesIO or None: If `output_path` is None, returns an in-memory BytesIO object containing the PDF.
      If `output_path` is specified, saves the PDF to the file and returns None.
    """
    if orientation == "landscape":
        page_width, page_height = landscape(A4)
    else:
        page_width, page_height = portrait(A4)

    # Set up the canvas to write to the provided output (file or buffer)
    if output_path:
        # Set up the canvas and page size based on orientation
        c = canvas.Canvas(output_path, pagesize=(page_width, page_height))
    else:
        output_buffer = io.BytesIO()
        c = canvas.Canvas(output_buffer, pagesize=(page_width, page_height))

    if image_paths is None:
        raise ValueError("image_paths must be provided and cannot be empty.")

    # Calculate the usable page area after accounting for the page margin
    usable_width = page_width - 2 * page_margin
    usable_height = page_height - 2 * page_margin

    # Calculate cell dimensions based on grid size and usable area
    cell_width = usable_width / columns
    cell_height = usable_height / rows

    # Loop through images and place them in the grid, handling multiple pages
    for i, image_path in enumerate(image_paths):
        # Determine the image"s aspect ratio
        img = ImageReader(image_path)
        init_img_width, init_img_height = img.getSize()
        image_aspect_ratio = init_img_width / init_img_height

        # Check if we need to create a new page (after filling the current page"s grid)
        if i > 0 and i % (columns * rows) == 0:
            c.showPage()  # Add a new page in the PDF

        if orientation == "auto" and (columns == 1 & rows == 1):
            if image_aspect_ratio > 1:
                page_width, page_height = landscape(A4)
                usable_width = page_width - 2 * page_margin
                usable_height = page_height - 2 * page_margin
                c.setPageSize((page_width, page_height))
            else:
                page_width, page_height = portrait(A4)
                usable_width = page_width - 2 * page_margin
                usable_height = page_height - 2 * page_margin
                c.setPageSize((page_width, page_height))

            c.drawImage(image_path, page_margin, page_margin, width=usable_width, height=usable_height,
                        preserveAspectRatio=True)
        else:
            # Calculate the grid position on the current page
            col = i % columns
            row = (i // columns) % rows

            # Calculate the x and y position for the image in its cell, accounting for page margin
            x = page_margin + col * cell_width + image_margin
            y = page_height - page_margin - ((row + 1) * cell_height) + image_margin  # Start from top of page

            # Calculate the width and height for the image within its cell, including image margins
            img_width = cell_width - 2 * image_margin
            img_height = cell_height - 2 * image_margin

            cell_aspect_ratio = img_width / img_height

            # Rotate the image if the aspect ratio suggests it will fit better rotated
            if (image_aspect_ratio > 1 and cell_aspect_ratio < 1) or (image_aspect_ratio < 1 and cell_aspect_ratio > 1):
                # Image is wider (landscape) and cell is taller (portrait) or vice versa -> rotate the image
                c.saveState()  # Save the canvas state to apply rotation
                # Rotate around the center of the image placement
                c.translate(x + img_width / 2, y + img_height / 2)
                c.rotate(90)
                # Adjust x, y since the image rotates around its center
                c.drawImage(image_path, -img_height / 2, -img_width / 2,
                            width=img_height, height=img_width, preserveAspectRatio=True)
                c.restoreState()  # Restore canvas state to avoid affecting other elements
            else:
                # Draw the image in the calculated position, fitting within the cell
                c.drawImage(image_path, x, y, width=img_width, height=img_height, preserveAspectRatio=True)

    # Save the PDF
    c.showPage()
    c.save()
    # If using an output buffer, return its contents to the start for reading
    if output_path:
        return output_path
    else:
        # If using an output buffer, return its contents to the start for reading
        output_buffer.seek(0)
        return output_buffer


def extract_and_merge_pdfs(pdf_paths, page_selections=None, output_pdf_path=None):
    """
    Extracts specific pages from multiple PDF files and combines them into a new PDF file.
    If page_selections is None or empty for a file, all pages from that file are included.

    Parameters:
    - pdf_paths: A list of paths to the source PDF files.
    - page_selections: A list of lists of tuples, where each inner list contains tuples representing
      individual pages or page ranges (as tuples) to be extracted from the corresponding PDF file in pdf_paths.
      If None or if any list inside is empty, all pages from that file will be included.
    - output_pdf_path (optional): The path for saving the output PDF to disk. If None, saves to an in-memory buffer.

    Returns:
    - If output_pdf_path is None, returns a BytesIO buffer containing the merged PDF.
    - If output_pdf_path is provided, saves the PDF to the specified path and returns None.
    """
    with pymupdf.open() as output_pdf:  # create a new PDF for the merged output
        for pdf_index, pdf_path in enumerate(pdf_paths):
            if not os.path.exists(pdf_path):
                print(f"File not found: {pdf_path}")
                continue

            # Open each PDF file with context manager
            with pymupdf.open(pdf_path) as pdf_document:
                last_page = pdf_document.page_count

                # Get the selection list for this PDF, or use an empty list if page_selections is None
                selections = page_selections[pdf_index] if page_selections and len(page_selections) > pdf_index else []
                if not selections:  # If selections is empty, add all pages
                    output_pdf.insert_pdf(pdf_document)
                else:
                    for selection in selections:
                        if len(selection) == 1:  # Single page tuple
                            page_num = selection[0] - 1  # Convert to 0-indexed
                            if 0 <= page_num < last_page:
                                output_pdf.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)
                        elif len(selection) == 2:  # Page range tuple
                            start_page, end_page = selection
                            start_page -= 1  # Convert to 0-indexed
                            end_page = min(end_page, last_page) - 1  # Adjust to last page if necessary
                            output_pdf.insert_pdf(pdf_document, from_page=start_page, to_page=end_page)

        # Save to a file if output_pdf_path is provided
        if output_pdf_path:
            output_pdf.save(output_pdf_path)
            print(f"Merged PDF created at: {output_pdf_path}")
            return output_pdf_path
        else:
            # Save to an in-memory buffer
            pdf_buffer = io.BytesIO()
            output_pdf.save(pdf_buffer)
            pdf_buffer.seek(0)  # Reset the buffer position to the start
            return pdf_buffer


if __name__ == "__main__":
    pdf_paths = ["file1.pdf", "file2.pdf", "file3.pdf"]
    page_selections = [[(1,)], [(1,)], [(1,)]]
    extract_and_merge_pdfs(pdf_paths, page_selections, "merged.pdf")
