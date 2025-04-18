from PIL import Image

# Open the image
image = Image.open("BE 1st.png")

# Show the image
image.show()

# Save the image to a file
image.save("certificate.png")
