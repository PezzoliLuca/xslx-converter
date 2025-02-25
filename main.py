from populate_h2020 import *
from populate_mic import *
from populate_mur import *
from populate_pnrr import *
import PySimpleGUI as sg
from openpyxl import load_workbook

def select_file(prompt_message):
    sg.theme('DarkBlue3')  # Tema opzionale
    file_path = sg.popup_get_file(prompt_message, file_types=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
    return file_path


def populate_template(template_file, output_file, amm_note_prog, mese_amm_prin, output_option):
    try:
        template_file_wb = load_workbook(template_file)
        if output_option == "H2020":
            print("H2020")
            populate_h2020(template_file_wb, output_file, amm_note_prog)

        elif output_option == "MIC":
            print("MIC")

            # Chiedere il range di mesi
            while True:
                try:
                    start_month = int(input(
                        "Inserisci il numero del mese di inizio (1 per gennaio, 2 per febbraio, ..., 12 per dicembre): "))
                    end_month = int(input(
                        "Inserisci il numero del mese di fine (1 per gennaio, 2 per febbraio, ..., 12 per dicembre): "))

                    if 1 <= start_month <= 12 and 1 <= end_month <= 12 and start_month <= end_month:
                        break
                    else:
                        print("Inserisci un range valido (esempio: 6 e 12 per Giugno - Dicembre).")
                except ValueError:
                    print("Input non valido. Inserisci numeri tra 1 e 12.")

            # Determina il numero di mesi selezionati
            num_months = end_month - start_month + 1

            # Se il range non è completo (1-12), carica un template specifico
            if not (start_month == 1 and end_month == 12):
                template_file = f"./templates/ESEMPIO_TS_MIC_{num_months}.xlsx"
                try:
                    template_file_wb = load_workbook(template_file)
                except FileNotFoundError:
                    print(f"Attenzione: Il template '{template_file}' non esiste. Verrà utilizzato quello di default.")

            populate_mic(template_file_wb, output_file, mese_amm_prin, start_month, end_month)

        elif output_option == "MUR":
            print("MUR")

            while True:
                try:
                    selected_month = int(
                        input("Inserisci il numero del mese (1 per gennaio, 2 per febbraio, ..., 12 per dicembre): "))
                    if 1 <= selected_month <= 12:
                        break
                    else:
                        print("Inserisci un valore compreso tra 1 e 12.")
                except ValueError:
                    print("Input non valido. Inserisci un numero tra 1 e 12.")

            month_duration = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
            n_days_in_month = month_duration[selected_month - 1]

            if n_days_in_month == 29:
                template_file = "./templates/ESEMPIO_TS_MUR_29gg.xlsx"
                template_file_wb = load_workbook(template_file)
            elif n_days_in_month == 30:
                template_file = "./templates/ESEMPIO_TS_MUR_30gg.xlsx"
                template_file_wb = load_workbook(template_file)

            populate_mur(template_file_wb, output_file, amm_note_prog, selected_month)

        elif output_option == "PNRR":
            print("PNRR")

            while True:
                try:
                    selected_month = int(
                        input("Inserisci il numero del mese (1 per gennaio, 2 per febbraio, ..., 12 per dicembre): "))
                    if 1 <= selected_month <= 12:
                        break
                    else:
                        print("Inserisci un valore compreso tra 1 e 12.")
                except ValueError:
                    print("Input non valido. Inserisci un numero tra 1 e 12.")

            month_duration = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
            n_days_in_month = month_duration[selected_month - 1]

            if n_days_in_month == 29:
                template_file = "./templates/ESEMPIO_TS_PNRR_29gg.xlsx"
                template_file_wb = load_workbook(template_file)
            elif n_days_in_month == 30:
                template_file = "./templates/ESEMPIO_TS_PNRR_30gg.xlsx"
                template_file_wb = load_workbook(template_file)

            populate_pnrr(template_file_wb, output_file, mese_amm_prin, selected_month)

        else:
            print("Opzione non riconosciuta.")

    except FileNotFoundError:
        print(f"Error: Template file '{template_file}' or input file not found.")
    except Exception as e:
        print(f"Error while processing template file '{template_file}': {e}")

def main():

    print("\nSelect an output option:")
    print("1. H2020")
    print("2. MIC")
    print("3. MUR")
    print("4. PNRR")

    option_map = {
        "1": "H2020",
        "2": "MIC",
        "3": "MUR",
        "4": "PNRR"
    }

    option = input("Enter the number corresponding to your choice: ")
    output_option = option_map.get(option)

    template_files = {
        "H2020": "./templates/ESEMPIO_TS_H2020.xlsx",
        "MIC": "./templates/ESEMPIO_TS_MIC_12.xlsx",
        "MUR": "./templates/ESEMPIO_TS_MUR_31gg.xlsx",
        "PNRR": "./templates/ESEMPIO_TS_PNRR_31gg.xlsx"
    }

    if output_option and output_option in template_files:
        template_file = template_files[output_option]
        output_file = f"output_{output_option.lower()}.xlsx"

        mese_amm_prin = amm_note_prog = None

        if output_option in ["H2020", "MUR"]:
            amm_note_prog = select_file("Select 'AMM_NOTE_PROG' Excel file")

        if output_option in ["MIC", "PNRR"]:
            mese_amm_prin = select_file("Select 'MESE_AMM_PRIN' Excel file")

        populate_template(template_file, output_file, amm_note_prog, mese_amm_prin, output_option)
    else:
        print("Invalid choice or missing template file. Exiting application.")

if __name__ == "__main__":
    main()
