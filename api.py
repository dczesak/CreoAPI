import PySimpleGUI as sg
import creopyson
import subprocess
import time
import os.path


try:
    c = creopyson.Client()
    c.connect()
except ConnectionError as e:
    print("Connection error")
    
if not c.is_creo_running():
    subprocess.Popen("D:/Creo/Creo 7.0.1.0/Parametric/bin/parametric.bat")
    time.sleep(30)

c.connect()
c.creo_set_creo_version(7)


dims_dict = {
    # [A, B, C, D, E, F]
    "W1": [100, 100, 8, 100, 36, 10],
    "W2": [100, 140, 8, 100, 36, 10],
    "W3": [100, 180, 8, 100, 36, 10],
    "W4": [140, 100, 10, 100, 38, 12],
    "W5": [140, 140, 10, 100, 38, 12],
    "W6": [140, 180, 10, 100, 38, 12],
    "W7": [180, 100, 11, 100, 40, 14],
    "W8": [180, 140, 11, 100, 40, 14],
    "W9": [180, 180, 11, 100, 40, 14]
}


def change_directory(path):
    c.file_close_window()
    c.creo_cd(path)
    c.file_open("model.prt")

def set_variant(variant):
    print(variant)
    A,B,C,D,E,F = dims_dict[variant]
    PatternA = 0
    if (A == 180):
        PatternA = 4
    elif (A == 140):
        PatternA = 3
    elif (A == 100):
        PatternA = 2
    
    PatternB = 0
    if (B == 180):
        PatternB = 4
    elif (B == 140):
        PatternB = 3
    elif (B == 100):
        PatternB = 2 
  
    c.feature_suppress(name="ELEMENT_A")
    c.feature_suppress(name="ELEMENT_B")
    c.parameter_set('A', A)
    c.parameter_set('B', B)
    c.parameter_set('C', C)
    c.parameter_set('D', D)
    c.parameter_set('E', E)
    c.parameter_set('F', F)
    c.parameter_set('PA', PatternA)
    c.parameter_set('PB', PatternB)
    if (PatternA > 2):
        c.feature_resume(name="ELEMENT_A")
    if (PatternB > 2):
        c.feature_resume(name="ELEMENT_B")
    c.file_regenerate()
    print(A,B,C,D,E,F)


def save_file(path, variant):
    A,B,C,D,E,F = dims_dict[variant]
    file = open(path, 'w')
    file.write("Nazwa: do_cz_13k1")
    file.write("\nA= {}".format(A))
    file.write("\nB= {}".format(B))
    file.write("\nC= {}".format(C))
    file.write("\nD= {}".format(D))
    file.write("\nE= {}".format(E))
    file.write("\nF= {}".format(F))
    file.write("\nMaterial name: {}".format(c.file_get_cur_material()))
    file.write("\nModel mass {}".format(str(c.file_massprops()['mass'])))
    file.close()
    print("Exported to text file to: " + path)
    

def export_pdf(path):
    c.interface_export_3dpdf(dirname=path)
    print("Exported to 3dpdf to: " + path)

def export_step(path):
    c.interface_export_file("STEP", dirname=path)
    print("Exported to STEP file to: " + path)

def set_material(material):
    c.file_set_cur_material(material)
    print("Material set to {}".format(c.file_get_cur_material()))
    

def insert_to_assemble():
    c.file_save()
    c.file_close_window()
    c.file_open("assemble.asm")
    c.file_assemble("model.prt")
    print("Inserted to assemble")


layout = [
    [sg.Input(size=(60,1), key='-DIRECTORY_INPUT-', enable_events=True), 
    sg.FolderBrowse('Choose folder', key='-DIRECTORY_BUTTON-', target='-DIRECTORY_INPUT-'), sg.Button('Confirm', key='-CONFIRM_DIRECTORY-')],
    [sg.Image('res/model_cad.png', expand_x=True, expand_y=True)],
    [
        sg.Frame(
            layout=[
                [
        sg.Frame(
            layout=[
                [
                    sg.Combo([
                    'W1','W2','W3','W4','W5','W6','W7','W8','W9'
                    ], 
                enable_events=True, key="-VARIANTS-", size=(8), default_value='W1', bind_return_key=True),
                sg.Text( "A: {}, B: {}, C: {}, D: {}, E: {}, F: {}".format(
                str(dims_dict['W1'][0]), str(dims_dict['W1'][1]), str(dims_dict['W1'][2]), str(dims_dict['W1'][3]), 
                str(dims_dict['W1'][4]), str(dims_dict['W1'][5]) 
            ), enable_events=True, key='-TEXT_VARIANTS-', size=(33)),
                sg.Button('Confirm', key='-VARIANT_CONFIRM-', size=(15))
                ]
            ],
            title="Variants"
        )
    ],
    [
        sg.Frame(
            layout=[
                [sg.Combo([
                'Cast_iron_ductile',
                'Steel_cast',
                'Steel_medium_carbon'
                ],
                size=(48,1), key='-MATERIAL_COMBO-', default_value='Cast_iron_ductile'), 
                sg.Button('Confirm', key='-MATERIAL_CONFIRM-', size=(15,1))]
            ],
            title="Material"
        )
    ],
    [
        sg.Frame(
            layout=[
                [sg.Input(size=(50,1), key='-FILE_INPUT-', default_text = os.getcwd()), 
                sg.FileSaveAs('Choose folder', key = '-FILE_SAVE_TXT-', size=(15,1), file_types=(('TXT', '.txt'),), enable_events=True),
                sg.Button('Save', key='-FILE_CONFIRM-', size=(10,1), target='-FILE_INPUT-')]
            ],
            title="Save to text file"
        )
    ],
    [
        sg.Frame(
            layout=[
                [sg.Input(size=(50,1), key='-FILE_STEP_INPUT-', default_text = os.getcwd()), 
                sg.FolderBrowse('Choose folder', size=(15,1)),
                sg.Button('Export', key='-STEP_CONFIRM-', size=(10,1), target='-FILE_STEP_INPUT-')]
            ],
            title="Export to STEP file"
        )
    ],
    [
        sg.Frame(
            layout=[
                [sg.Input(size=(50,1), key='-FILE_3DPDF_INPUT-', default_text = os.getcwd()), 
                sg.FolderBrowse('Choose folder', size=(15,1)),
                sg.Button('Export', key='-3DPDF_CONFIRM-', size=(10,1), target='-FILE_3DPDF_INPUT-')]
            ],
            title="Export to 3D PDF file"
        )
    ],
    
        [
        sg.Frame(
            layout=[
                [sg.Button('Insert to assemble', key='-INSERT_ASSEMBLE-', size=(60,1))]
            ],
            title="Inserting to assemble", 
        )
        ],      
            ],
            title="Options"
        )
    ]
]

sg.theme('DarkTanBlue')
window = sg.Window('API', layout, resizable=True, element_justification='c')

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    if event == '-VARIANTS-':
        variant = values['-VARIANTS-']
        window['-TEXT_VARIANTS-'].update(
            "A: {}, B: {}, C: {}, D: {}, E: {}, F: {}".format(
                str(dims_dict[variant][0]), str(dims_dict[variant][1]), str(dims_dict[variant][2]), str(dims_dict[variant][3]), 
                str(dims_dict[variant][4]), str(dims_dict[variant][5]) 
            )
        )
    if event == '-CONFIRM_DIRECTORY-':
        path = str(values['-DIRECTORY_INPUT-'])
        print(path)
        change_directory(path)
    if event == '-VARIANT_CONFIRM-':
        variant = values['-VARIANTS-']
        set_variant(variant)
    if event == '-FILE_CONFIRM-':
        path = values['-FILE_INPUT-']
        variant = values['-VARIANTS-']
        save_file(path, variant)
    if event == '-MATERIAL_CONFIRM-':
        material = values['-MATERIAL_COMBO-']
        set_material(material)
    if event == '-STEP_CONFIRM-':
        path = values['-FILE_STEP_INPUT-']
        export_step(path)
    if event == '-3DPDF_CONFIRM-':
        path = values['-FILE_3DPDF_INPUT-']
        export_pdf(path)
    if event == '-INSERT_ASSEMBLE-':
        insert_to_assemble()
    if event == '-FILE_SAVE_TXT-':
        name = values['-FILE_SAVE_TXT-']
        print(name)
window.close()