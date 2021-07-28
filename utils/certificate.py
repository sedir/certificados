import os
import traceback
from io import BytesIO
from pathlib import Path

import matplotlib.font_manager as fontman
from PIL import Image

from utils import image_utils


def find_font_file(query):
    matches = list(filter(lambda path: query in os.path.basename(path), fontman.findSystemFonts()))
    return matches


def scale_to_width(dimensions, width):
    height = (width * dimensions[1]) / dimensions[0]
    return int(width), int(height)


def open_image(path):
    im = Image.open(path)

    pixel_data = im.load()

    if im.mode == "RGBA":
        # If the image has an alpha channel, convert it to white
        # Otherwise we'll get weird pixels
        for y in range(im.size[1]):  # For each row ...
            for x in range(im.size[0]):  # Iterate through each column ...
                # Check if it's opaque
                if pixel_data[x, y][3] < 255:
                    # Replace the pixel data with the colour white
                    pixel_data[x, y] = (255, 255, 255, 255)
    return im


def generate_certificate(text, certificate: Image or Path or str, pos_x, pos_y, width, font_size, line_spacing,
                         alignment, color, font, data=None, preview=True):
    # TODO: there's gotta be a better way!
    font_path = font

    if isinstance(certificate, (str, Path)):
        im = open_image(certificate)
    else:
        im = certificate

    text_y = im.height * pos_y / 100
    text_x = im.width * pos_x / 100
    text_width = im.width * width / 100

    im = image_utils.ImageText(im)

    if data:
        try:
            text = text.format(*data.values())
        except Exception as e:
            traceback.print_exc()

    im.write_text_box(text_x, text_y, box_width=text_width, font_filename=str(font_path),
                      text=text, font_size=font_size, line_spacing=line_spacing, color=color, place=alignment)

    if preview:
        size = scale_to_width(im.image.size, 400)
        b = BytesIO()
        im.image.resize(size, Image.BICUBIC).convert('RGB').save(b, format='PNG')
        return b.getvalue(), size
    else:
        output = BytesIO()
        im.image.convert("RGB").save(output, format='PDF')
        output.seek(0)
        return output
