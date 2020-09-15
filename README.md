# pptx_image_slides
A tool for automatically generating PowerPoint presentations with with image-only slides

# Purpose:
Suppose you have a folder full of images and you want to create a PowerPoint presentation in which each image is rescaled to take up the entire area of one slide. 
This project allows you to do this, with the option to append these image-slides to an existing PowerPoint presentation, or to create a new PowerPoint presentation with the image-slides.


# Usage:
There is just one function, `create_slides()`, that controls the whole process.

```
import pptx_image_slides as slides

img_folder_path = 'test_data'

pptx_path = 'test_data/example.pptx'

slides.create_slides(img_folder_path, pptx_path)
```

This will create the PowerPoint file with all the image-slides

# Documentation:
```
def create_slides(img_folder_path,
                  pptx_path, 
                  slide_height_inch=7.5, 
                  slide_width_inch=13.333,
                  extensions=('.png', '.jpg', '.tif')):

This function can create image-based slides and add them to a power point presentation in one of
three modes.
1) Create: If the given presentation path does not exist, a new presentation is created
2) Copy and Append: If the given presentation path does exist, a copy is made, new slides are appended, and
the presentation is saved with the original filename plus an '_edit' suffix.

img_folder_path: a string path to the images folder

pptx_path: a string path for the PowerPoint file

slide_height_inch: the height of the slide in inches

slide_width_inch: the width of the slide in inches

extensions: a list of image extensions strings like ['.png', '.jpg']
```
