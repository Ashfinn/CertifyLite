from pptx import Presentation
import os
import comtypes.client  # Required for PDF conversion (Windows only)

def generate_certificates(template_path, names, output_folder):
    # Load the PowerPoint template
    prs = Presentation(template_path)
    
    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)
    
    ppt_files = []  # Store PPTX file paths

    for name in names:
        # Create a copy of the template
        new_prs = Presentation(template_path)
        
        for slide in new_prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "Shanta Das" in run.text:  # Replace the name
                                run.text = run.text.replace("Shanta Das", name)
        
        # Save the new certificate
        pptx_path = os.path.join(output_folder, f"Certificate_{name}.pptx")
        new_prs.save(pptx_path)
        ppt_files.append(pptx_path)
        print(f"Certificate saved: {pptx_path}")

    return ppt_files

def convert_ppt_to_pdf(ppt_files, output_folder):
    """ Converts PPTX files to PDFs using PowerPoint COM object (Windows) """
    os.makedirs(output_folder, exist_ok=True)
    
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1  # Make it visible (optional)

    for ppt_path in ppt_files:
        pdf_path = os.path.splitext(ppt_path)[0] + ".pdf"
        presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
        presentation.SaveAs(pdf_path, 32)  # 32 = PDF format
        presentation.Close()
        print(f"PDF saved: {pdf_path}")

    powerpoint.Quit()

# Example usage
names_list = ["Obidur Rahman", "Bob Smith", "Charlie Brown"]  # Replace with your names
pptx_template = "cert.pptx"  # Your file in the directory
output_directory = "Certificates"

ppt_files = generate_certificates(pptx_template, names_list, output_directory)
convert_ppt_to_pdf(ppt_files, output_directory)
