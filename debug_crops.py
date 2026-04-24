# Save this as debug_crops.py in your RFQ_automation folder
from PIL import Image
import io
Image.MAX_IMAGE_PIXELS = None

img = Image.open("rfq_inputs/RFQ_VW_001/tz___2fk_863_021____1_v1_0.tif")
W, H = img.size
print(f"Original: {W}x{H}")

num_strips = 5
for i in range(num_strips):
    crop = img.crop((int(W*i/num_strips), 0, int(W*(i+1)/num_strips), H))
    crop.thumbnail((4500, 4500), Image.LANCZOS)
    crop.save(f"output/debug_strip_{i+1}.jpg", quality=95)
    print(f"Strip {i+1} saved: {crop.size}")

print("Done — check output/ folder")