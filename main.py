import pandas as pd
import os


def main():
    #Função principal, todas as funçoes são chamadas para atender ao que é nescessario
    sitelist_file = "SiteList.xlsx"
    results_file = "Results.xlsx"
    report_file = "result.xlsx"

    df_sitelist = read_data(sitelist_file)
    df_results = read_data(results_file)


    if df_sitelist is not None and df_results is not None:
        #Verifica se as planilhas não estão vazias antes de gerar o merge
        df_final = generate_report(df_sitelist, df_results)
        save_report(df_final, report_file)

    #Chamando as funçoes para a implementação
    quality_0(df_final)
    mbps_more_80(df_final)
    mbps_less_10(df_final)
    sites_not_in_results(df_results, df_final, 'Site ID')

def read_data(file_name):
    #Função para ler os arquivos
    pd.read_excel(file_name)
    try:
        df = pd.read_excel(file_name)
        return df
    except pd.errors.ParserError as e:
        print(f"Error reading {file_name}: {e}")
        return None

def generate_report(df_sitelist,df_results):
    #Função para realizar o merge e gerar o novo DF
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
    #Utilizando a biblioteca OS para salvar o excel na mesma pasta da main
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_file_path = os.path.join(script_dir, file_name)
    df.to_excel(excel_file_path, index=False)

def alert_yes(df):
    #Realizando a querry para procurar o que é pedido
    print("Sites com velocidade Menor que 10 Mbps:")
    df_alerta_yes = df.query('Alert == "Yes"')
    if df_alerta_yes.empty:
        print("Nenhum site encontrado com alerta ativo.")
        print('--------------------------------------')
    else:
        print(df_alerta_yes.to_string())
        print('--------------------------------------')

def quality_0(df):
    print("Sites com qualidade = 0:")
    df_qualidade = df.query('`Quality (0-10)` == 0')
    if df_qualidade.empty:
        print("Nenhum site encontrado com qualidade = 0.")
        print('--------------------------------------')
    else:
        print(df_qualidade.to_string())
        print('--------------------------------------')

def mbps_more_80(df):
    print("Site com velocidade maior que 80 Mbps:")
    df_maior_80 = df.query('Mbps > 80')
    if df_maior_80.empty:
        print("Nenhum site encontrado com velocidade maior do que 80.")
        print('--------------------------------------')
    else:
        print(df_maior_80.to_string())
        print('--------------------------------------')

def mbps_less_10(df):
    print("Sites com velocidade Menor que 10 Mbps:")
    df_menor_10 = df.query('Mbps < 10')
    if df_menor_10.empty:
        print("Nenhum site encontrado com velocidade menor do que 10.")
        print('--------------------------------------')
    else:
        print(df_menor_10.to_string())
        print('--------------------------------------')

def sites_not_in_results(results_df, other_df, column_name):
    #busca quais são os sites que se encontram em uma planilha e não na outra
    print("Sites que não estão presentes no Results:")

    results_sites = set(results_df[column_name])
    other_sites = set(other_df[column_name])

    sites_not_in_results = other_sites - results_sites

    if not sites_not_in_results:
        print("Todos os sites estão presentes no Results.")
    else:
        print("Sites ausentes no Results:")
        for site in sites_not_in_results:
            print(site)

if __name__ == '__main__':
    main()




