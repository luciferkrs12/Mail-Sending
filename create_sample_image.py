from PIL import Image, ImageDraw, ImageFont
import os

def create_sample_image(image_path, output_path, name="Sample Name"):
    try:
        img = Image.open(image_path)
        draw = ImageDraw.Draw(img)
        width, height = img.size

        
        # --- PATCHING LOGIC TO REMOVE "NAME" ---
        # Strategy: Tile a clean section of the ribbon to cover the center area.
        # This prevents the blurry look of stretching.
        
        y_ribbon_start = int(height * 0.69)
        y_ribbon_end = int(height * 0.77)
        ribbon_height = y_ribbon_end - y_ribbon_start
        
        # Target area to cover (Center where "NAME" is)
        target_x_start = int(width * 0.38) # Narrowed slightly to be safe
        target_x_end = int(width * 0.62)
        target_width = target_x_end - target_x_start
        
        # Source area (Clean ribbon on left)
        src_x_start = int(width * 0.20)
        src_x_end = int(width * 0.28) # Take a smaller, safer clean chunk
        src_width = src_x_end - src_x_start
        
        clean_slice = img.crop((src_x_start, y_ribbon_start, src_x_end, y_ribbon_end))
        
        # Tile the slice to cover the target width
        current_x = target_x_start
        while current_x < target_x_end:
            paste_width = min(src_width, target_x_end - current_x)
            if paste_width < src_width:
                # Crop the last piece if needed
                patch = clean_slice.crop((0, 0, paste_width, ribbon_height))
            else:
                patch = clean_slice
            
            img.paste(patch, (current_x, y_ribbon_start))
            current_x += src_width
        
        # --- DRAWING TEXT ---
        
        name = name.upper() # FORCE UPPERCASE
        
        # Color: Dark Maroon
        text_color = (60, 0, 0) 
        
        # Max width for text 
        max_text_width = int(width * 0.55) 
        
        # Dynamic Scaler
        current_font_size = int(width * 0.06) # Start bigger
        min_font_size = int(width * 0.03)
        
        # Load Font - Switch to Serif (Times New Roman) for premium look
        font_names = ["timesbd.ttf", "georgiab.ttf", "arialbd.ttf"]
        font_path = None
        for fn in font_names:
            possible_path = f"C:/Windows/Fonts/{fn}"
            if os.path.exists(possible_path):
                font_path = possible_path
                break
        
        if not font_path:
             font_path = "C:/Windows/Fonts/arial.ttf"

        while current_font_size > min_font_size:
            try:
                font = ImageFont.truetype(font_path, current_font_size)
            except OSError:
                font = ImageFont.load_default()
                break

            if hasattr(draw, "textbbox"):
                bbox = draw.textbbox((0, 0), name, font=font)
                text_w = bbox[2] - bbox[0]
                text_h = bbox[3] - bbox[1]
            else:
                text_w, text_h = draw.textsize(name, font=font)
                
            if text_w <= max_text_width:
                break 
            
            current_font_size -= 2
        
        # Calculate centered position
        x = (width - text_w) / 2
        
        # Center vertically in the ribbon patch
        ribbon_middle = y_ribbon_start + (ribbon_height / 2)
        # Adjust vertical center slightly for font baseline
        y = ribbon_middle - (text_h / 2) - (text_h * 0.15) 

        # Draw text
        draw.text((x, y), name, font=font, fill=text_color)
        
        img.save(output_path)
        print(f"Sample image saved to {output_path}")

    except Exception as e:
        print(f"Error creating sample image: {e}")

if __name__ == "__main__":
    create_sample_image("Congratulations.png", "sample_output.png", "VIJAY SURYA")
