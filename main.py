import os

import matplotlib.font_manager as fontman
import pandas as pd
import PySimpleGUI as sg

from utils import certificate

font_list = sorted(fontman.findSystemFonts(fontpaths=None, fontext='ttf'))

col_left = [[sg.Text('Certificate template'), sg.Input(key='path_model'),
             sg.FileBrowse('search', target='path_model',
                           file_types=(('Imagens JPG', '*.jpg'), ('Imagens PNG', '*.png'),))],
            [sg.Text('Input spreadsheet'), sg.Input(key='path_sheet'),
             sg.FileBrowse('search', target='path_sheet')],
            [sg.Text('Certificate text'), sg.Multiline(key='certificate_text'),
             sg.ColorChooserButton('Text color', key='color')],
            [sg.Text('Font'), sg.Combo(font_list, key='font', readonly=True)],
            [sg.Text('Font size'), sg.Slider(range=(12, 172), orientation='h', size=(48, 20),
                                                    change_submits=True, key='slider_font')],
            [sg.Text('Line spacing'), sg.Slider(range=(1, 50), orientation='h', size=(44, 20),
                                                            default_value=10.0,
                                                            change_submits=True, key='slider_linespacing')],
            [sg.Text('Alignment'), sg.Radio('Left', "alignment", key='align_left'),
             sg.Radio('Right', "alignment", key='align_right'),
             sg.Radio('Center', "alignment", key='align_center'),
             sg.Radio('Justify', "alignment", key='align_justify', default=True), ],
            [sg.Text('Position X (%)'), sg.Slider(range=(1, 100), orientation='h', size=(50, 20), resolution=0.5,
                                                 change_submits=False, default_value=15.0, key='slider_x')],
            [sg.Text('Position Y (%)'), sg.Slider(range=(1, 100), orientation='h', size=(50, 20), resolution=0.5,
                                                 change_submits=False, default_value=35.0, key='slider_y')],
            [sg.Text('Width (%)'), sg.Slider(range=(1, 100), orientation='h', default_value=70.0,
                                               size=(51, 20), resolution=0.5,
                                               change_submits=False, key='width')],
            [sg.Text('Output folder'), sg.Input(key='path_output'),
             sg.FolderBrowse('search', target='path_output')],
            ]

col_right = [[sg.Text('Preview:')],
             [sg.Image(data='', size=(400, 309), key='preview')]]

layout = [
    [sg.Column(col_left), sg.Column(col_right)],
    [sg.OK(), sg.Cancel('Cancel')],
]

window = sg.Window("Certificate generator").Layout(layout)
# Event Loop

current_values = None
cache = {}

while True:
    event, values = window.Read(timeout=0)
    if current_values is None:
        current_values = values
    if event is None:
        break

    path = values['path_model']
    path_data = values['path_sheet']
    path_output = values['path_output']

    alignment = [x[0].split('_')[1] for x in values.items() if 'align_' in x[0] and x[1] is True][0]
    color = values['color'].lstrip('#') if values['color'] else '000000'
    color = tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))

    if path and os.path.exists(path) and not os.path.isdir(path) and current_values != values:
        current_values = values
        if path not in cache.keys():
            cache[path] = certificate.open_image(path)

        example_data = None
        if path_data and os.path.exists(path_data) and not os.path.isdir(path_data) and path_data not in cache.keys():
            if path_data.lower().endswith('.xlsx'):
                cache[path_data] = pd.read_excel(path_data).to_dict('records')
            elif path_data.lower().endswith('.csv'):
                cache[path_data] = pd.read_csv(path_data).to_dict('records')

        if path_data in cache.keys():
            example_data = cache[path_data][0]

        im, size = certificate.generate_certificate(values['certificate_text'], cache[path].copy(), values['slider_x'],
                                                    values['slider_y'], values['width'], int(values['slider_font']),
                                                    int(values['slider_linespacing']), alignment, color, values['font'],
                                                    example_data, True)
        window.FindElement('preview').Update(data=im, size=size)

    if event == 'Generate':

        if path_data and os.path.exists(path_data) and not os.path.isdir(
                path_data) and path_output and os.path.exists(path_output) and os.path.isdir(path_output):
            for i, x in enumerate(cache[path_data]):
                sg.OneLineProgressMeter('Processing certificates', i + 1, len(cache[path_data]), 'meter',
                                        'Please wait...')
                cert_data = certificate.generate_certificate(values['certificate_text'], cache[path].copy(),
                                                             values['slider_x'],
                                                             values['slider_y'], values['width'],
                                                             int(values['slider_font']),
                                                             int(values['slider_linespacing']), alignment, color,
                                                             values['font'],
                                                             x, False)
                with open(os.path.join(path_output, "{:04d}.pdf".format(i)), 'wb') as f:
                    f.write(cert_data.read())
