# This code example demonstrates how to convert DWF to PDF with options
import aspose.cad as cad

# Load DWF file
image = cad.Image.load("C://dwf//A12.030276.idw.dwf")

# Specify CAD Rasterization options
rasterizationOptions = cad.imageoptions.CadRasterizationOptions()

# Specify PDF options
pdfOptions = cad.imageoptions.PdfOptions()
pdfOptions.vector_rasterization_options = rasterizationOptions
rasterizationOptions.draw_type = cad.fileformats.cad.CadDrawTypeMode.USE_OBJECT_COLOR

# Save as PDF
image.save("C://dwf//A12.030276.pdf", pdfOptions)
