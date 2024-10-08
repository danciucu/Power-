import os
from pptx import Presentation

def extract_images_and_captions_from_ppt(pptx_path, output_folder):
    # Load the presentation
    prs = Presentation(pptx_path)
    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.mkdir(output_folder)
    # Initialize image counter
    image_counter = 1
    # Array to hold image file paths and captions
    image_caption_dictionary = {}

    # Loop through slides
    for slide_num, slide in enumerate(prs.slides):
        # Loop through shapes in each slide
        for shape in slide.shapes:
            # Check if the shape contains a picture
            if hasattr(shape, "image"):
                # Get the image in the shape
                image = shape.image
                # Define output image name
                image_name = f"image_{slide_num + 1}.jpg"
                # Define output image path
                image_path = os.path.join(output_folder, image_name)

                # Save the image
                with open(image_path, "wb") as img_file:
                    img_file.write(image.blob)

                # Initialize caption as None (in case no caption is found)
                caption = None

                # Try to find the caption for this image (assumes it's in a nearby shape)
                for other_shape in slide.shapes:
                    if other_shape != shape and other_shape.has_text_frame:
                        # You can adjust this condition to better detect captions, based on your slides' structure
                        # Here, we assume that if the other shape is positioned below the image, it's the caption
                        if other_shape.top > shape.top:  # Caption is usually below the image
                            caption = other_shape.text.strip()
                            break  # Assume the first matching text box is the caption

                # Append the image path and caption to the array
                image_caption_dictionary[image_path] = caption
                #print(f"Extracted: {image_path} with caption: {caption if caption else 'No Caption'}")
                image_counter += 1

    return image_caption_dictionary