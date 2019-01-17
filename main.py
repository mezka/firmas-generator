
import xlrd
import base64

fn_html_template = "./firma.html"
fn_xls = "./empleados.xlsx"

image_dictionary = [
    {
        "src": "./mesquita.png",
        "style": "display:block;height:100px;margin: 12px 0;",
        "alt": "mesquita-logo"
    },
    {
        "src": "./iram_2015.png",
        "style": "display:block;height:100px;margin: 12px 0;",
        "alt": "iram-logo"
    }
]

def load_image_dictionary_with_base64_images():
    for image in image_dictionary:
        with open(image["src"], "rb") as image_file:
            image["base64"] = base64.b64encode(image_file.read()).decode("utf-8")


def load_employee_dictionaries():

    book = xlrd.open_workbook(fn_xls)
    sh = book.sheet_by_index(0)

    employee_dictionaries = []
    dColumnToKey = {}

    for cx in range(sh.ncols):
        dColumnToKey[cx] = sh.cell_value(rowx=0, colx=cx)

    for rx in range(sh.nrows):
        employee_info = {}
        for cx in range(sh.ncols):
            employee_info[dColumnToKey[cx]] = sh.cell_value(rowx=rx, colx=cx)

        employee_dictionaries.append(employee_info)

    employee_dictionaries.pop(0) #the first entry on the list is the first row of the xls, which is the keys to the dictionary

    return employee_dictionaries

def write_out():

    employee_dictionaries = load_employee_dictionaries()
    load_image_dictionary_with_base64_images()

    for employee_dictionary in employee_dictionaries:
        string_out = string_in
        for key in employee_dictionary.keys():
            string_out = string_out.replace(key, employee_dictionary[key])
            fn_out = f'./out/Firma_{employee_dictionary["#NOMBRE#"]}.html'.replace(' ', '_')
        with open(fn_out, "w+") as f_out:
            image_string_out = ""
            for idx, image in enumerate(image_dictionary):
                image_string_out += f'<img src="data:image/png;filename=file{idx};base64,{image["base64"]}" style="{image["style"]}" alt="{image["alt"]}"></img>\n'
            
            string_out = string_out.replace("#IMAGENES#", image_string_out)
            f_out.write(string_out)

with open(fn_html_template) as f_template:
    string_in = f_template.read()

write_out()

