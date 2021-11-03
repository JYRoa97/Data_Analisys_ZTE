import pandas as pd
import func_filtro as ff
def print_hi(name):
    df_cn = pd.read_excel("C:/Users/Admin/OneDrive/ProyectoMinTik7K/29_oct/CN_Maestro.xlsx")
    df_sm = pd.read_excel("C:/Users/Admin/OneDrive/ProyectoMinTik7K/29_oct/CARLOS DIAZ.xlsx")

    df_cn["cod"] = df_cn["Site"].str.extract(r"(\d{5})")
    df_sm = df_sm.rename(columns={"ID Beneficario": "cod"})
    df_sm = df_sm[df_sm['cod'].notna()]

    df_sm["cod"] = df_sm["cod"].apply(str)

    filtro = df_cn[df_cn.cod.isin(df_sm.cod)]
    # filtro["IM"] = filtro.groupby(by=["cod", "Device Name"])["Status"].apply(list)
    print(filtro)
    df_status = filtro.groupby(by=["cod", "Device Name"])["Status"].value_counts().unstack().fillna(0)
    print(df_status)

    df_status.to_excel("C:/Users/Admin\OneDrive/ProyectoMinTik7K/APsStatus-29-10-2021.xlsx")


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #print_hi('PyCharm')
    ff.main()
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
