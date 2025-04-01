import win32com.client as win32
import pandas as pd
import time

def OpenHysys(hysys_file_path):
    print("Opening HYSYS...")
    hysysApp = win32.Dispatch("HYSYS.Application")
    hysysCase = hysysApp.SimulationCases.Open(hysys_file_path)
    hysysCase.Visible = True
    print("HYSYS file opened.")
    return hysysCase

def GetInitialConditions(hysysCase, mode):
    spreadsheet = hysysCase.Flowsheet.Operations.Item("SPRDSHT-1")
    try:
        base_data = {
            "Mole Fraction of Nitrogen (Reactor-feed) [-]": spreadsheet.Cell("D2").CellValue,
            "Mole Fraction of Nitrogen (R1-out) [-]": spreadsheet.Cell("D4").CellValue,
            "Mole Fraction of Hydrogen (Reactor-feed) [-]": spreadsheet.Cell("F2").CellValue,
            "Mole Fraction of Hydrogen (R1-out) [-]": spreadsheet.Cell("F4").CellValue,
            "Mole Fraction of Ammonia (Reactor-feed) [-]": spreadsheet.Cell("E2").CellValue,
        }
        
        if mode == "pressure":
            base_data.update({
                "Reactor-feed temperature(°C)": spreadsheet.Cell("B2").CellValue,
                "temprature of (R1-out)(°C)": spreadsheet.Cell("B4").CellValue,
                "Ammonia inlet temperature(°C)": spreadsheet.Cell("B3").CellValue,
            })
        elif mode == "temperature":
            base_data.update({
                "Reactor-feed pressure(kPa)": spreadsheet.Cell("C2").CellValue,
                "Ammonia inlet pressure(kPa)": spreadsheet.Cell("C3").CellValue,
            })
        
        return base_data
    except Exception as e:
        print(f"Error fetching initial conditions: {e}")
        return None

def SaveToExcel(sheet_name, base_data, results):
    excel_path = "hysys_output.xlsx"
    base_df = pd.DataFrame(list(base_data.items()), columns=["Parameter", "Value"])
    results_df = pd.DataFrame(results)
    
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        base_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)
        results_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=base_df.shape[0] + 3)
    
    print(f"Output saved to {excel_path} in sheet {sheet_name}")

def UpdatePressure(hysysCase):
    reactor_feed = hysysCase.Flowsheet.MaterialStreams.Item("Reactor-feed")
    ammonia = hysysCase.Flowsheet.MaterialStreams.Item("ammonia")
    spreadsheet = hysysCase.Flowsheet.Operations.Item("SPRDSHT-1")
    
    base_data = GetInitialConditions(hysysCase, "pressure")
    if not base_data:
        return
    
    start_pressure = float(input("Enter starting pressure (kPa): "))
    end_pressure = float(input("Enter ending pressure (kPa): "))
    interval = float(input("Enter interval for pressure (kPa): "))
    
    sheet_name = f"Pressure_{int(start_pressure)}_{int(end_pressure)}"
    results = []
    
    for new_pressure in range(int(start_pressure), int(end_pressure) + 1, int(interval)):
        reactor_feed.Pressure = new_pressure
        ammonia.Pressure = new_pressure
        time.sleep(1)
        
        try:
            mole_fraction_ammonia = spreadsheet.Cell("E4").CellValue
        except Exception as e:
            print(f"Error fetching ammonia mole fraction: {e}")
            mole_fraction_ammonia = None
        
        results.append({
            "Reactor-feed Pressure [kPa]": new_pressure,
            "Ammonia Pressure [kPa]": new_pressure,
            "Mole Fraction of Ammonia in R1-out [-]": mole_fraction_ammonia
        })
    
    SaveToExcel(sheet_name, base_data, results)

def UpdateTemperature(hysysCase):
    reactor_feed = hysysCase.Flowsheet.MaterialStreams.Item("Reactor-feed")
    ammonia = hysysCase.Flowsheet.MaterialStreams.Item("ammonia")
    spreadsheet = hysysCase.Flowsheet.Operations.Item("SPRDSHT-1")
    
    base_data = GetInitialConditions(hysysCase, "temperature")
    if not base_data:
        return
    
    start_temp = float(input("Enter starting temperature (°C): "))
    end_temp = float(input("Enter ending temperature (°C): "))
    interval = float(input("Enter interval for temperature (°C): "))
    
    sheet_name = f"Temperature_{int(start_temp)}_{int(end_temp)}"
    results = []
    
    for new_temp in range(int(start_temp), int(end_temp) + 1, int(interval)):
        reactor_feed.Temperature = new_temp
        ammonia.Temperature = new_temp
        time.sleep(1)
        
        try:
            mole_fraction_ammonia = spreadsheet.Cell("E4").CellValue
        except Exception as e:
            print(f"Error fetching ammonia mole fraction: {e}")
            mole_fraction_ammonia = None
        
        results.append({
            "Reactor-feed Temperature [°C]": new_temp,
            "Ammonia Temperature [°C]": new_temp,
            "Mole Fraction of Ammonia in R1-out [-]": mole_fraction_ammonia
        })
    
    SaveToExcel(sheet_name, base_data, results)

hysys_file_path = r"C:\Users\cclab1\Desktop\swapn\swapn_pc.hsc"
hysysCase = OpenHysys(hysys_file_path)

choice = input("Do you want to vary Pressure (P) or Temperature (T)? ").strip().lower()
if choice == 'p':
    UpdatePressure(hysysCase)
elif choice == 't':
    UpdateTemperature(hysysCase)
else:
    print("Invalid choice. Please enter 'P' for Pressure or 'T' for Temperature.")
