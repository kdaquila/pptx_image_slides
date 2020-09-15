from unittest import TestCase
import pptx_image_slides as slides


class Test(TestCase):
    def test_modify_path(self):
        original_path = "C:\\folder1\\folder2\\myFile.txt"
        modified_path = slides.modify_path(original_path)
        expected_path = "C:\\folder1\\folder2\\myFile_edit.txt"
        self.assertEqual(expected_path, modified_path)

    def test_find_images(self):
        actual_img_paths = slides.find_images("..\\test_data")
        expected_img_paths = ['..\\test_data\\Image1.png',
                              '..\\test_data\\Image2.png',
                              '..\\test_data\\Image3.png']
        self.assertEqual(expected_img_paths, actual_img_paths)

    def test_create_blank_prs(self):
        prs = slides.create_blank_prs(7, 8)
        actual_slide_size = (prs.slide_height.inches, prs.slide_width.inches)
        expected_slide_size = (7, 8)
        self.assertEqual(expected_slide_size, actual_slide_size)

    def test_add_image_slide(self):
        prs = slides.create_blank_prs()
        slides.add_image_slide(prs, "..\\test_data\\Image1.png")
        actual_num_slides = len(prs.slides)
        expected_num_slides = 1
        self.assertEqual(expected_num_slides, actual_num_slides)
