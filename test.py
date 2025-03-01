# from PIL import Image

# # Open the transparent overlay image (this will be the reference size)
# overlay = Image.open("Test(2).png").convert("RGBA")

# # Open the background image
# background = Image.open("image.png").convert("RGBA")

# # Resize the background to match the overlay's size
# background = background.resize(overlay.size)

# # Blend the images
# combined = Image.alpha_composite(background, overlay)

# # Show or save the result
# combined.show()  # Show the final image
# combined.save("output.png")  # Save the final image

from PIL import Image, ImageEnhance

# Open the transparent overlay image (reference size)
overlay = Image.open("Test(1).png").convert("RGBA")

# Open the background image
background = Image.open("image.png").convert("RGBA")

# Resize background to match the overlay's size
background = background.resize(overlay.size)

# Increase brightness (adjust factor as needed)

# Adjust transparency of the background
alpha_value = 175  # Range: 0 (fully transparent) to 255 (fully opaque)
overlay.putalpha(alpha_value)

# Blend the transparent brightened background with the overlay
combined = Image.alpha_composite(Image.new("RGBA", overlay.size, (0, 0, 0, 0)), background)
combined = Image.alpha_composite(combined, overlay)

# enhancer = ImageEnhance.Brightness(combined)
# combined= enhancer.enhance(0.5)  # Increase brightness (1.0 = original, >1.0 = brighter)
# Show or save the result
combined.show()  # Display the final image
combined.save("output.png")  # Save the image
