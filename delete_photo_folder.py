import os
import glob

def delete_images(folder_path):
    # Get a list of all image files in the folder
    image_files = glob.glob(os.path.join(folder_path, '*.jpg')) + glob.glob(os.path.join(folder_path, '*.jpeg')) + glob.glob(os.path.join(folder_path, '*.png')) + glob.glob(os.path.join(folder_path, '*.gif')) + glob.glob(os.path.join(folder_path, '*.bmp'))

    # Iterate over the list and delete each file
    for image_file in image_files:
        os.remove(image_file)
        print(f"Deleted: {image_file}")


folder_path = "qr_img"
delete_images(folder_path)
