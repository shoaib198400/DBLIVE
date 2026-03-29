from PIL import Image

# Convert PNG to ICO
img = Image.open(r"D:/SHOAIB/VS CODE PROJECTS/EXCEPTION SNAPSHOT DASHBOARD/MAster/Desktopicon.png")
img.save(r"D:/SHOAIB/VS CODE PROJECTS/EXCEPTION SNAPSHOT DASHBOARD/MAster/Desktopicon.ico", format="ICO", sizes=[(256,256)])
print("Icon conversion complete.")
