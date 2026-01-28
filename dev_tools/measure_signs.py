from PIL import Image
import os

img_path = 'extracted_assets/signs.png'
if not os.path.exists(img_path):
    print(f"Error: {img_path} not found")
    exit(1)

img = Image.open(img_path)
width, height = img.size
print(f"Image dimensions: {width} x {height}")

# Convert to grayscale to simplify finding signatures
gray = img.convert('L')
# Invert colors so signatures are light (if they were dark)
# or just look for non-white pixels.
# Let's look for horizontal "centers" of gravity for 3 clusters.

# Sample the image horizontally to find signature "islands"
cols = []
for x in range(width):
    # Check if column has dark pixels (threshold 200)
    has_ink = False
    for y in range(height):
        if gray.getpixel((x, y)) < 230:
            has_ink = True
            break
    cols.append(has_ink)

# Find clusters of has_ink
clusters = []
in_cluster = False
start = 0
for i, has_ink in enumerate(cols):
    if has_ink and not in_cluster:
        start = i
        in_cluster = True
    elif not has_ink and in_cluster:
        clusters.append((start, i))
        in_cluster = False
if in_cluster:
    clusters.append((start, width))

print(f"Found {len(clusters)} ink clusters:")
for i, (s, e) in enumerate(clusters):
    center = (s + e) / 2.0
    print(f"Cluster {i+1}: {s} to {e} (Center: {center}, {center/width*100:.1f}%)")
