import PySimpleGUI as sg
from fpdf import FPDF

# Create the template
template = """
[addressee]
[company name]
[P.O.Box] 
[Location] 

[SUBJECT]
The Office acknowledges receipt of your letter dated [date letter was received] on the above subject.
We have reviewed the submitted documents taken notice of in your application. The Office will revert to you when your services are required.

We thank you for the interest expressed in the Energy sector of Ghana.

Yours faithfully,

[DIGITAL SIGNATURE]

[NAME OF OFFICER]
[PORTFOLIO]

"""

# Create the input form layout
layout = [
    [sg.Text("Addressee"), sg.InputText()],
    [sg.Text("Company name"), sg.InputText()],
    [sg.Text("P.O.Box"), sg.InputText()],
    [sg.Text("Location"), sg.InputText()],
    [sg.Text("Subject"), sg.InputText()],
    [sg.Text("Date letter was received"), sg.InputText()],
    [sg.Submit(), sg.Cancel()]
]

# Create the form window
form = sg.Window("Fill in the template").Layout(layout)

# Loop until the form is submitted
while True:
    button, values = form.Read()
    if button == "Submit":
        # Replace the placeholders in the template with the input values
        filled_template = template.replace("[addressee]", values[0])
        filled_template = filled_template.replace("[company name]", values[1])
        filled_template = filled_template.replace("[P.O.Box]", values[2])
        filled_template = filled_template.replace("[Location]", values[3])
        filled_template = filled_template.replace("[SUBJECT]", values[4])
        filled_template = filled_template.replace("[date letter was received]", values[5])
        # Create the PDF object
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Times", 12)
        pdf.multi_cell(0, 10, filled_template)
        pdf.output("C:/Users/lesli/Desktop/HAULAGE_GENERATED PDFs/output.pdf")
        sg.popup("Template filled and exported to PDF.")
        break
    if button == "Cancel":
        break

