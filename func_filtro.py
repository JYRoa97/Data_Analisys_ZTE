import pandas as pd

def main():
    excel_cn_maestro = pd.ExcelFile("C:/Users/Admin/OneDrive/ProyectoMinTik7K/01-nov/CN_Maestro.xlsx")
    d_CnMaestro = pd.read_excel(excel_cn_maestro, "CN Maestro")
    excel_SM = pd.ExcelFile("C:/Users/Admin/OneDrive/ProyectoMinTik7K/01-nov/CARLOS DIAZ.xlsx")
    d_SM = pd.read_excel(excel_SM, "IM")
    pd.set_option('max_rows', 5)
    print("Setup complete.")

    array_delete = d_CnMaestro.loc[d_CnMaestro['Site'] == '777-PILOTO'].index.to_numpy()
    d_CnMaestro = d_CnMaestro.drop(array_delete)
    #print(d_CnMaestro)
    d_CnMaestro["cod"] = d_CnMaestro["Site"].str.extract(r"(\d{5})")
    d_SM = d_SM.rename(columns={"ID Beneficiario": "ID_Beneficiario"})
    #print(d_CnMaestro)

    d_SM = d_SM[d_SM.ID_Beneficiario.notna()]

    d_SM['ID_Beneficiario'] = d_SM['ID_Beneficiario'].apply(int)
    #print(d_SM)
    d_SM['ID_Beneficiario'] = d_SM['ID_Beneficiario'].apply(str)
    #print(d_SM)
    filtro = d_CnMaestro[d_CnMaestro.cod.isin(d_SM.ID_Beneficiario)]
    #print(filtro)
    # ------------------------------------------- GENERAR EXCEL -----------------------------------------------------------------

    d_Aps = filtro.groupby(['cod', "Device Name"])['Status'].value_counts().unstack().fillna(0)
    d_Aps_1 = d_Aps
    #print(d_Aps_1)
    #_____________________________________________________________________________________________________________________________
    d_Aps = filtro.groupby(['cod'])['Status'].value_counts().unstack().fillna(0)

    d_Aps_visitar_CDs = d_Aps.loc[(d_Aps['Online'] > 0.0)]
    d_Aps_visitar_CDs = d_Aps_visitar_CDs.loc[(d_Aps_visitar_CDs['Online'] < 3.0)]
    d_Aps_visitar_conectante = d_Aps.loc[d_Aps['Online'] == 0.0]
    #print(d_Aps_visitar_CDs)


    d_Aps = d_Aps.rename_axis('cod').reset_index()
    d_Aps.index.name = None
    d_Aps_visitar_CDs = d_Aps_visitar_CDs.rename_axis('cod').reset_index()
    d_Aps_visitar_CDs.index.name = None
    d_Aps_visitar_conectante = d_Aps_visitar_conectante.rename_axis('cod').reset_index()
    d_Aps_visitar_conectante.index.name = None

    # ------------------------------------------- GENERAR EXCEL -----------------------------------------------------------------
    d_Aps_visitar_conectante.drop(columns=['Offline', 'Online'])
    # ___________________________________________________________________________________________________________________________
    d_Aps_visitar_CDs.drop(columns=['Offline', 'Online'])
    d_Aps = d_Aps.rename(columns={"cod": "ID_Beneficiario"})
    d_Aps_visitar_CDs = d_Aps_visitar_CDs.rename(columns={"cod": "ID_Beneficiario"})
    d_Aps_visitar_conectante = d_Aps_visitar_conectante.rename(columns={"cod": "ID_Beneficiario"})
    #filtro_dos = d_CnMaestro[d_CnMaestro.cod.isin(d_Aps_visitar_CDs.ID_Beneficiario)]

    #print(filtro_dos)
    #d_Aps_visitar_CDs = filtro_dos.groupby(['cod', "Device Name"])['Status'].value_counts().unstack().fillna(0)
    #print(d_Aps_visitar_CDs)
    #filtro_tres = d_CnMaestro[d_CnMaestro.cod.isin(d_Aps_visitar_conectante.ID_Beneficiario)]

    #print(filtro_tres)
    #d_Aps_visitar_conectante = filtro_tres.groupby(['cod', "Device Name"])['Status'].value_counts().unstack().fillna(0)
    #print(d_Aps_visitar_conectante)
    to_excel_sheet(d_Aps_1,d_Aps_visitar_conectante,d_Aps_visitar_CDs)



def to_excel_sheet(df_Aps,df_Tx,df_CDs):
    print("Excel creado con éxito")
    with pd.ExcelWriter("C:/Users/Admin/OneDrive\ProyectoMinTik7K/01-nov/APsStatus-01-11-2021.xlsx") as writer:
        df_Aps.to_excel(writer, sheet_name="Aps_Status")
        df_Tx.to_excel(writer, sheet_name="Aps_Revisar_TX")
        df_CDs.to_excel(writer, sheet_name="Aps_Revisar_CDs")
