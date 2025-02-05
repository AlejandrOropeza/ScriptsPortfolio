import pandas as pd

ruta_original = r"C:\Users\orope\Documents\Documentos digitalizados\Proceso Wallmart\Base_Nielsen_Prueba.xlsx"
ruta_procesado = ruta_original.replace(".xlsx", "_Procesado.xlsx")

xls = pd.ExcelFile(ruta_original)

hojas_originales = {sheet: pd.read_excel(xls, sheet_name=sheet, header=None) for sheet in xls.sheet_names}

df_comercial = pd.read_excel(xls, sheet_name="Extracción áreas", skiprows=10)
df_marca_propia = pd.read_excel(xls, sheet_name="Extracción áreas MP", skiprows=10)

df_comercial.columns = df_comercial.columns.astype(str)
df_marca_propia.columns = df_marca_propia.columns.astype(str)

df_marca_propia["Tipo_Producto"] = "CONTROLLED LABEL"
df_comercial["Tipo_Producto"] = "MARCA COMERCIAL"

df_base_nielsen = pd.concat([df_comercial, df_marca_propia], ignore_index=True)

df_base_nielsen.columns = df_base_nielsen.columns.astype(str)

col_mercado = next((col for col in df_base_nielsen.columns if "mercado" in col.lower()), None)

if col_mercado:
    df_base_nielsen.rename(columns={col_mercado: "MERCADO"}, inplace=True)
else:
    raise KeyError("No se encontró la columna 'MERCADO' en los datos. Revisa el archivo original.")

df_base_nielsen["Formato"] = df_base_nielsen["MERCADO"].str[:-2]
df_base_nielsen["Región"] = df_base_nielsen["MERCADO"].str[-2:]

region_map = {
    "A1": "PACIFICO",
    "A2": "NORTE",
    "A3": "OCCIDENTE",
    "A4": "CENTRO",
    "A5": "VDM",
    "A6": "SURESTE"
}

df_base_nielsen["Región"] = df_base_nielsen["Región"].map(region_map)

with pd.ExcelWriter(ruta_procesado, engine="openpyxl") as writer:
    for sheet, df in hojas_originales.items():
        df.to_excel(writer, sheet_name=sheet, index=False, header=False)

    df_base_nielsen.to_excel(writer, sheet_name="Base_Nielsen", index=False, header=True)

print(f"Proceso completado. Archivo guardado en: {ruta_procesado}")