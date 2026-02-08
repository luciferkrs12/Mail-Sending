from PIL import Image

def crop_logo(input_path, output_path):
    try:
        img = Image.open(input_path)
        bbox = img.getbbox()
        if bbox:
            cropped_img = img.crop(bbox)
            cropped_img.save(output_path, optimize=True)
            print(f"Successfully cropped {input_path} to {output_path}. New size: {cropped_img.size}")
        else:
            print(f"No content found in {input_path}")
            
    except Exception as e:
        print(f"Error processing image: {e}")

if __name__ == "__main__":
    crop_logo("sm_logo_small.png", "sm_logo_small.png")
