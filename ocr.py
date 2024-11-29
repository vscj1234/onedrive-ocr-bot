import cv2
import numpy as np
import pytesseract
from PIL import Image
from io import BytesIO
import os

class OCRProcessor:
    def __init__(self):
        pytesseract.pytesseract.tesseract_cmd = r"E:\Tesseract-OCR\tesseract.exe" # Change path to your local tesseract application
        self.supported_formats = {'.png', '.jpg', '.jpeg', '.tiff', '.bmp'}

    def preprocess_image(self, image_bytes: bytes) -> Image.Image:
        """Preprocess the image for better OCR results."""
        try:
            # First try with PIL
            image = Image.open(BytesIO(image_bytes))
            
            # Convert to RGB if necessary
            if image.mode != 'RGB':
                image = image.convert('RGB')
            
            # Convert to numpy array
            img_array = np.array(image)
            
            # Convert to grayscale
            gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
            
            # Apply adaptive thresholding
            binary = cv2.adaptiveThreshold(
                gray, 
                255, 
                cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                cv2.THRESH_BINARY, 
                11, 
                2
            )
            
            # Denoise
            denoised = cv2.fastNlMeansDenoising(binary)
            
            # Scale the image (upscale if too small)
            height, width = denoised.shape
            if height < 1000 or width < 1000:
                scale_factor = 2
                denoised = cv2.resize(
                    denoised, 
                    None, 
                    fx=scale_factor, 
                    fy=scale_factor, 
                    interpolation=cv2.INTER_CUBIC
                )
            
            return Image.fromarray(denoised)
            
        except Exception as e:
            print(f"Image preprocessing error: {str(e)}")
            return None

    def extract_text(self, image_bytes: bytes) -> str:
        """Extract text from image using OCR."""
        try:
            # Preprocess the image
            processed_image = self.preprocess_image(image_bytes)
            if processed_image is None:
                print("Failed to preprocess image")
                return ""
            
            # Save processed image for debugging (optional)
            # processed_image.save('processed_image.png')
            
            # Configure OCR parameters
            custom_config = r'--oem 3 --psm 3 -l eng --dpi 300'
            
            # Perform OCR
            text = pytesseract.image_to_string(
                processed_image,
                config=custom_config
            )
            
            # Clean the extracted text
            cleaned_text = ' '.join(text.split())
            
            if not cleaned_text:
                print("No text extracted. Trying alternative preprocessing...")
                # Try alternative preprocessing
                img_array = np.array(Image.open(BytesIO(image_bytes)).convert('RGB'))
                gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
                # Increase contrast
                contrast = cv2.convertScaleAbs(gray, alpha=1.5, beta=0)
                # Try OCR again
                text = pytesseract.image_to_string(
                    contrast,
                    config=custom_config
                )
                cleaned_text = ' '.join(text.split())
            
            print(f"Extracted text length: {len(cleaned_text)}")
            print(f"First 100 characters: {cleaned_text[:100]}")
            
            return cleaned_text if cleaned_text else ""
            
        except Exception as e:
            print(f"OCR processing error: {str(e)}")
            print(f"Current Tesseract path: {pytesseract.pytesseract.tesseract_cmd}")
            return ""

    def is_image_file(self, filename: str) -> bool:
        """Check if the file is a supported image format."""
        return any(filename.lower().endswith(fmt) for fmt in self.supported_formats)
