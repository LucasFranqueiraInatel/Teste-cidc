import pandas as pd
import os


def main():
    sitelist_file = "SiteList.xlsx"
    results_file = "Results.xlsx"
    report_file = "result.xlsx"

    df_sitelist = read_data(sitelist_file)
    df_results = read_data(results_file)


    if df_sitelist is not None and df_results is not None:
        df_final = generate_report(df_sitelist, df_results)
        save_report(df_final, report_file)  # report_file should be 'result.xlsx' if that's the intended file name

    df_final = read_data(report_file)

    qualidade_0(df_final)
    maior_80(df_final)
    menor_10(df_final)
    sites_not_in_results(df_results, df_final, 'Site ID')

def read_data(file_name):
    pd.read_excel(file_name)
    try:
        df = pd.read_excel(file_name)
        return df
    except pd.errors.ParserError as e:
        print(f"Error reading {file_name}: {e}")
        return None

def generate_report(df_sitelist,df_results):
    df_sitelist_columns = df_sitelist[["Site Name", "Site ID"]]
    df_results_columns = df_results[["Site ID", "Equipment","Signal (%)","Quality (0-10)","Mbps","Year"]]
    df = df_sitelist_columns.merge(df_results_columns, on="Site ID")


    df = df.assign(
        Site=df["Site Name"].str.slice(9).str[:-2],
        State=df["Site Name"].str[-2:]

    )

    df = df.drop("Site Name", axis=1)
    df = df.query("Year == 2023")

    df_relatorio = df[[
        "Site",
        "Site ID",
        "Equipment",
        "State",
        "Signal (%)",
        "Quality (0-10)",
        "Mbps",
    ]]
    return df_relatorio

def save_report(df, file_name):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_file_path = os.path.join(script_dir, file_name)
    df.to_excel(excel_file_path, index=False)

def qualidade_0(df):
    print("Sites com qualidade = 0:")
    df_qualidade = df.query('`Quality (0-10)` == 0')
    if df_qualidade.empty:
        print("Nenhum site encontrado com qualidade = 0.")
        print('--------------------------------------')
    else:
        print(df_qualidade.to_string())
        print('--------------------------------------')

def maior_80(df):
    print("Site com velocidade maior que 80 Mbps:")
    df_maior_80 = df.query('Mbps > 80')
    if df_maior_80.empty:
        print("Nenhum site encontrado com velocidade maior do que 80.")
        print('--------------------------------------')
    else:
        print(df_maior_80.to_string())
        print('--------------------------------------')

def menor_10(df):
    print("Sites com velocidade Menor que 10 Mbps:")
    df_menor_10 = df.query('Mbps < 10')
    if df_menor_10.empty:
        print("Nenhum site encontrado com velocidade menor do que 10.")
        print('--------------------------------------')
    else:
        print(df_menor_10.to_string())
        print('--------------------------------------')

def sites_not_in_results(results_df, other_df, column_name):
    print("Sites que não estão presentes no Results:")

    # Get unique site names from both DataFrames
    results_sites = set(results_df[column_name])
    other_sites = set(other_df[column_name])

    # Sites that are in other_df but not in results_df
    sites_not_in_results = other_sites - results_sites

    if not sites_not_in_results:
        print("Todos os sites estão presentes no Results.")
    else:
        print("Sites ausentes no Results:")
        for site in sites_not_in_results:
            print(site)

if __name__ == '__main__':
    main()












# print("Sites com alerta ativos:")

#def alerta_ativo():
# df_alertas = df_busca.query('Alerts == "Yes"')
# # print(df_alertas.to_string())

#def qualidade_0():
#   print("Sites com 0 de qualidade:")
#   df_qualidade = df_busca.query('`Quality (0-10)` == 0')
#   print(df_qualidade.to_string())
#
#def maior_80():
#   df_maior_80 = df_busca.query('Mbps > 80')
#   print(df_maior_80.to_string())

#def menor_10():
#   df_menor_10 = df_busca.query('Mbps < 10')
#   print(df_menor_10.to_string())

