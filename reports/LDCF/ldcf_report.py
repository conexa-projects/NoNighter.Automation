import os
import time
import psutil
from pywinauto import Application, timings
import win32com.client
import pyautogui

# Inicializar Excel con win32com
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True

# Abrir archivo de Excel
desktop_path = os.path.join(os.environ['USERPROFILE'], 'OneDrive', 'Desktop')
file_path = os.path.join(desktop_path, 'Test MVP - Template.xlsx')
workbook = excel.Workbooks.Open(file_path)

# Esperar a que Excel se cargue completamente
time.sleep(3)

# Obtener el PID del proceso de Excel
excel_pid = None
for proc in psutil.process_iter(['pid', 'name']):
    if proc.info['name'] == 'EXCEL.EXE':
        excel_pid = proc.info['pid']
        break

if excel_pid is None:
    print("No se encontró el proceso de Excel.")
else:
    print(f"Conectando a Excel con PID: {excel_pid}")

    app = Application(backend="uia").connect(process=excel_pid)
    dlg = app.window(title_re=".*Excel")

    # Restaurar y maximizar la ventana si está minimizada o no maximizada
    if dlg.is_minimized():
        dlg.restore()
    dlg.maximize()

    # Excel en primer plano
    dlg.set_focus()
    time.sleep(3)

    # Hacer click en el Add-In usando el nombre del Add-In
    addin_button = dlg.child_window(title="NoNighter QA", control_type="Button")
    addin_button.click_input()
    time.sleep(3)

    pyautogui.click(x=1700, y=400)
    time.sleep(3)
    pyautogui.click(x=1700, y=530)

    # Ingresar nombre al reporte
    input_field = dlg.child_window(auto_id="valuation-name", control_type="Edit")
    timings.wait_until_passes(10, 0.5, lambda: input_field.is_enabled())
    input_field.set_focus()
    time.sleep(1)
    input_field.type_keys("LDCF Automated Valuation", with_spaces=True)
    print("Nombre del reporte ingresado correctamente.")

    # Hacer click en el primer botón encontrado
    all_buttons = dlg.descendants(control_type="Button")
    if all_buttons:
        all_buttons[0].click_input()
        print("Click en el primer botón encontrado como alternativa.")
    else:
        print("No se encontraron botones para hacer click.")

    # Diccionario con inputs, valores y toggles para el Step 1
    step_1_data = {
        "checkbox-UDCF": ("checkbox", None),
        "checkbox-LDCF": ("checkbox", None),
        "VD": ("input_text", "12312027"),
        "FCFY": ("input_text", "2037"),
        "COE": ("input_text", "18.0"),
        "checkbox-PGRC": ("checkbox", None),
        "TEM": ("input_text", "4.0")
    }

    # Diccionario para el Step 2
    step_2_data = {
        "YS": ("input_text", "2025;2026;2027;2028;2029;2030;2031;2032;2033;2034;2035;2036;2037"),
        "EBI": ("input_text",
                "15,000,000.0;15,000,000.0;35,000,000.0;50,000,000.0;80,000,000.0;15,000,000.0;15,000,000.0;5,000,000.0;70,000,000.0;90,000,000.0;5,000,000.0;15,500,000.0;2,500,000,000.0"),
        "D&A": ("input_text",
                "2,100,000.0;2,982,000.0;4,234,440.0;6,012,904.8;8,536,324.8;12,124,421.2;17,216,678.2;24,447,683.0;34,715,709.8;49,296,308.0;70,000,757.3;99,401,075.4;141,149,527.1"),
        "CX": ("input_text",
               "4,400,000.0;4,532,000.0;4,667,960.0;4,807,998.8;4,952,238.8;5,100,805.9;8,000,000.0;8,240,000.0;8,487,200.0;8,741,816.0;9,004,070.0;9,274,192.6;9,552,418.4"),
        "CNWC": ("input_text",
                 "5,000,000.0;4,000,000.0;9,000,000.0;5,000,000.0;1,000,000.0;3,000,000.0;2,500,000.0;3,100,000.0;5,000,000.0;5,000,000.0;5,000,000.0;5,000,000.0;5,000,000.0"),
        "NI": ("input_text", "9,000,000.0;12,000,000.0;21,000,000.0;30,000,000.0;48,000,000.0;9,000,000.0;0.0;3,000,000.0;42,000,000.0;54,000,000.0;3,000,000.0;9,300,000.0;1,500,000,000.0"),
        "DIR": ("input_text", "-2,040,000.0;-2,040,000.0;-2,040,000.0;-2,040,000.0;-2,040,000.0;5,100,000.0;-2,040,000.0;-2,040,000.0;5,712,000.0;-2,040,000.0;8,160,000.0;-2,040,000.0;-2,040,000.0"),
        "button-add-new-cash-flow-item": ("button", None),
        "Cash Flow Item 0": ("input_text", "1,200,564.0;1,284,603.6;1,374,525.7;1,470,742.5;1,573,694.5;1,683,853.1;1,801,722.8;1,927,843.4;2,062,792.5;2,207,187.9;2,361,187.9;2,361,691.1;2,527,009.5;2,703,900.1"),
        "CCE": ("input_text", "20,000,000.0"),
        "TD": ("input_text", "100,000,000.0"),
        "MI": ("input_text", "5,000,000.0")
    }

    # Diccionario para el Step 3
    step_3_data = {
        "checkbox-PEM": ("checkbox", None),
        "NI": ("input_text", "9,000,000.0;12,000,000.0;21,000,000.0;30,000,000.0;48,000,000.0;9,000,000.0;0.0;3,000,000.0;42,000,000.0;54,000,000.0;3,000,000.0;9,300,000.0;1,500,000,000.0"),
        "COES": ("input_text", "1"),
        "TEMS": ("input_text", "0.5"),
    }

    # Llenar inputs del Step 1
    for auto_id, (element_type, value) in step_1_data.items():
        try:
            if element_type == "input_text":
                input_field = dlg.child_window(auto_id=auto_id, control_type="Edit")
                timings.wait_until_passes(10, 0.5, lambda: input_field.is_enabled())
                input_field.set_focus()
                time.sleep(1)
                input_field.type_keys(value, with_spaces=True)
                print(f"Texto '{value}' ingresado en el input con Automation ID '{auto_id}'.")

            elif element_type == "checkbox":
                checkbox = dlg.child_window(auto_id=auto_id, control_type="CheckBox")
                timings.wait_until_passes(10, 0.5, lambda: checkbox.is_enabled())
                if checkbox.get_toggle_state() == 1:  # Si está seleccionado
                    checkbox.click_input()  # Desmarcar el checkbox
                else:
                    checkbox.click_input()
        except Exception as e:
            print(f"Error al interactuar con el elemento '{auto_id}': {str(e)}")

    # Click en el botón "Next" para pasar al Step 2
    try:
        btn_next = dlg.child_window(auto_id="button-next", control_type="Button")
        timings.wait_until_passes(10, 0.5, lambda: btn_next.is_enabled())
        btn_next.click_input()
        time.sleep(3)
        print("Avanzado al Step 2.")
    except Exception as e:
        print(f"Error al hacer click en 'Next': {str(e)}")

    # Llenar inputs del Step 2
    for auto_id, (element_type, value) in step_2_data.items():
        try:
            if element_type == "input_text":
                input_field = dlg.child_window(auto_id=auto_id, control_type="Edit")
                timings.wait_until_passes(10, 0.5, lambda: input_field.is_enabled())
                input_field.set_focus()
                time.sleep(1)
                input_field.type_keys(value, with_spaces=True)
                print(f"Texto '{value}' ingresado en el input con Automation ID '{auto_id}'.")
            elif element_type == "button":
                add_new_cf_item_button = dlg.child_window(auto_id=auto_id, control_type="Button")
                timings.wait_until_passes(10, 0.5, lambda: add_new_cf_item_button.is_enabled())
                add_new_cf_item_button.click_input()

        except Exception as e:
            print(f"Error al interactuar con el elemento '{auto_id}': {str(e)}")

    # Click en el botón "Next" para avanzar al Step 3
    try:
        next_button_step2 = dlg.child_window(auto_id="button-next", control_type="Button")
        timings.wait_until_passes(10, 0.5, lambda: next_button_step2.is_enabled())
        next_button_step2.click_input()
        print("Botón 'Next' clickeado para avanzar al Step 3.")
    except Exception as e:
        print(f"Error al hacer clic en el botón 'Next' en Step 2: {str(e)}")

    # Interactuar con los inputs del Step 3
    for auto_id, (element_type, value) in step_3_data.items():
        try:
            if element_type == "checkbox":
                checkbox = dlg.child_window(auto_id=auto_id, control_type="CheckBox")
                timings.wait_until_passes(10, 0.5, lambda: checkbox.is_enabled())
                if checkbox.get_toggle_state() == 1:  # Si está seleccionado
                    checkbox.click_input()  # Desmarcar el checkbox
                print(f"Checkbox Present Value of Terminal Value / Enterprise Value desmarcado.")

            elif element_type == "input_text":
                input_field = dlg.child_window(auto_id=auto_id, control_type="Edit")
                timings.wait_until_passes(10, 0.5, lambda: input_field.is_enabled())
                input_field.set_focus()
                input_field.select().type_keys("{BACKSPACE}")  # Seleccionar y borrar el valor existente
                time.sleep(0.5)
                input_field.type_keys(value, with_spaces=True)  # Ingresar el nuevo valor
                print(f"Texto '{value}' ingresado en el input con Automation ID '{auto_id}'.")

        except Exception as e:
            print(f"Error al interactuar con el elemento '{auto_id}': {str(e)}")

    # Click en el botón "Finish" para generar el reporte
    try:
        finish_button = dlg.child_window(title="Finish", control_type="Button")
        timings.wait_until_passes(10, 0.5, lambda: finish_button.is_enabled())
        finish_button.click_input()
        print("Reporte generado correctamente.")
    except Exception as e:
        print(f"Error al hacer click en el botón 'Finish': {str(e)}")

# Cerrar Excel
# workbook.Close(SaveChanges=False)
# excel.Quit()
