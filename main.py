import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

def find_section(data_sheet, keyword):
    for idx, row in data_sheet.iterrows():
        if row.astype(str).str.contains(keyword, case=False, na=False).any():
            return idx
    return None

def extract_matrix(data_sheet, start_row, num_rows, num_cols, start_col=0):
    return data_sheet.iloc[start_row:start_row+num_rows, start_col:start_col+num_cols]

def clean_dataframe(df):
    df = df.dropna(how='all', axis=0)
    df = df.dropna(how='all', axis=1)
    return df.reset_index(drop=True)

def read_excel_data(file_path):
    data_sheet = pd.read_excel(file_path, header=None)

    prop_start_row = find_section(data_sheet, "PROPIEDADES MECÁNICAS DE LA ESTRUCTURA")
    if prop_start_row is not None:
        mechanical_properties = extract_matrix(data_sheet, prop_start_row+2, 5, 5)
        mechanical_properties = clean_dataframe(mechanical_properties)
    else:
        mechanical_properties = None

    coord_start_row = find_section(data_sheet, "COORDENADAS DEL SISTEMA ESTRUCTURAL")
    if coord_start_row is not None:
        nodal_coordinates = extract_matrix(data_sheet, coord_start_row+2, 4, 7)
        nodal_coordinates = clean_dataframe(nodal_coordinates)
    else:
        nodal_coordinates = None

    phys_start_row = find_section(data_sheet, "PROPIEDADES FÍSICAS DE LOS ELEMENTOS")
    if phys_start_row is not None:
        physical_properties = extract_matrix(data_sheet, phys_start_row+2, 4, 8)
        physical_properties = clean_dataframe(physical_properties)
    else:
        physical_properties = None

    stiffness_start_row = find_section(data_sheet, "MATRIZ DE RIGIDEZ DEL ELEMENTO FLEXIBLE")
    if stiffness_start_row is not None:
        stiffness_matrix = extract_matrix(data_sheet, stiffness_start_row+2, 12, 12)
        stiffness_matrix = clean_dataframe(stiffness_matrix)
    else:
        stiffness_matrix = None

    return mechanical_properties, nodal_coordinates, physical_properties, stiffness_matrix

def calculate_displacements_and_reactions(stiffness_global, forces):
    K_ff = stiffness_global[:3, :3]
    F_f = forces[:3]
    displacements_free = np.linalg.solve(K_ff, F_f)
    K_cf = stiffness_global[3:, :3]
    F_c = forces[3:]
    reactions = np.dot(K_cf, displacements_free) + F_c
    return displacements_free, reactions

def write_results_to_excel(output_file, results):
    with pd.ExcelWriter(output_file, engine='openpyxl', mode='w') as writer:
        for name, df in results.items():
            if df is not None:
                df.to_excel(writer, sheet_name=name, index=False)

def plot_and_save_diagrams(axial_force, shear_force, moment, x, output_directory="output"):
    os.makedirs(output_directory, exist_ok=True)
    plt.figure(figsize=(8, 5))
    plt.plot(x, axial_force, label="Fuerza Axial")
    plt.title("Diagrama de Fuerza Axial")
    plt.xlabel("Posición a lo largo del elemento (m)")
    plt.ylabel("Fuerza Axial (kN)")
    plt.legend()
    plt.grid(True)
    plt.savefig(os.path.join(output_directory, "fuerza_axial.png"))
    plt.close()
    plt.figure(figsize=(8, 5))
    plt.plot(x, shear_force, label="Fuerza Cortante", color="orange")
    plt.title("Diagrama de Fuerza Cortante")
    plt.xlabel("Posición a lo largo del elemento (m)")
    plt.ylabel("Fuerza Cortante (kN)")
    plt.legend()
    plt.grid(True)
    plt.savefig(os.path.join(output_directory, "fuerza_cortante.png"))
    plt.close()
    plt.figure(figsize=(8, 5))
    plt.plot(x, moment, label="Momento Flector", color="green")
    plt.title("Diagrama de Momento Flector")
    plt.xlabel("Posición a lo largo del elemento (m)")
    plt.ylabel("Momento Flector (kN·m)")
    plt.legend()
    plt.grid(True)
    plt.savefig(os.path.join(output_directory, "momento_flector.png"))
    plt.close()

if __name__ == "__main__":
    file_path = "SEGUNDOEXAMENBRAZOSRIGIDOS3D2.xlsx"
    output_file = "Resultados_Sistema_Estructural.xlsx"
    
    mechanical_properties, nodal_coordinates, physical_properties, stiffness_matrix = read_excel_data(file_path)

    stiffness_global = np.eye(6) * 1000
    forces = np.array([100, 0, -200, 0, 0, 0])
    displacements, reactions = calculate_displacements_and_reactions(stiffness_global, forces)

    results = {
        "Propiedades Mecánicas": mechanical_properties,
        "Coordenadas de los Nodos": nodal_coordinates,
        "Propiedades Físicas": physical_properties,
        "Matriz de Rigidez Local": stiffness_matrix,
        "Desplazamientos": pd.DataFrame(displacements, columns=["Desplazamientos"]),
        "Reacciones": pd.DataFrame(reactions, columns=["Reacciones"])
    }
    write_results_to_excel(output_file, results)

    x = np.linspace(0, 10, 100)
    axial_force = -100 * x
    shear_force = 50 * (10 - x)
    moment = 0.5 * x * (10 - x)
    plot_and_save_diagrams(axial_force, shear_force, moment, x)