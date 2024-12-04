import os
import subprocess
import zipfile
import rarfile
import pandas as pd
from tkinter import Tk, filedialog, messagebox
from tkinter import Button

# Função para descompactar o arquivo baseado no tipo (zip ou rar)
def descompactar_arquivo(rar_file_path, extracted_folder_path):
    if rar_file_path.endswith('.zip'):
        with zipfile.ZipFile(rar_file_path, 'r') as zip_ref:
            zip_ref.extractall(extracted_folder_path)
    elif rar_file_path.endswith('.rar'):
        # Usando WinRAR para descompactar arquivos .rar
        command = f'winrar x "{rar_file_path}" "{extracted_folder_path}\\*" -y'
        subprocess.run(command, shell=True)
    else:
        raise ValueError("Formato de arquivo não suportado. Somente '.zip' e '.rar' são permitidos.")

# Função para listar o conteúdo da pasta de extração
def listar_conteudo_da_pasta(extracted_folder_path):
    print("Conteúdo da pasta de extração:")
    for arquivo in os.listdir(extracted_folder_path):
        print(arquivo)

# Função para ler os arquivos CSV
def ler_arquivos_csv(extracted_folder_path):
    listar_conteudo_da_pasta(extracted_folder_path)  # Verifique o conteúdo da pasta

    # Verificando se o arquivo correto existe
    clientes_csv_path = os.path.join(extracted_folder_path, 'v_clientes_CodEmpresa_92577.csv')
    processos_csv_path = os.path.join(extracted_folder_path, 'v_processos_CodEmpresa_92577.csv')

    # Verificar se os arquivos existem
    if not os.path.exists(clientes_csv_path):
        raise FileNotFoundError(f"Arquivo não encontrado: {clientes_csv_path}")
    if not os.path.exists(processos_csv_path):
        raise FileNotFoundError(f"Arquivo não encontrado: {processos_csv_path}")

    # Ler os arquivos CSV
    clientes_df = pd.read_csv(clientes_csv_path, encoding='ISO-8859-1', delimiter=';')
    processos_df = pd.read_csv(processos_csv_path, encoding='ISO-8859-1', delimiter=';')
    
    return clientes_df, processos_df

# Função para tratar e unir os dados
def tratar_dados(clientes_df, processos_df):
    # Unir as tabelas com base no campo id_cliente (ajustar se necessário)
    merged_df = pd.merge(clientes_df, processos_df, left_on='id_cliente', right_on='id_cliente', how='inner')

    # Padronização de datas para o formato DD/MM/AAAA
    merged_df['data'] = pd.to_datetime(merged_df['data'], errors='coerce').dt.strftime('%d/%m/%Y')

    # Preencher valores nulos com 'Desconhecido' (ajustável)
    merged_df.fillna('Desconhecido', inplace=True)

    return merged_df

# Função para salvar os dados tratados em um arquivo CSV
def salvar_arquivo_migrado(merged_df, output_file_path):
    merged_df.to_csv(output_file_path, index=False)
    return output_file_path

# Função para carregar o arquivo de backup e executar a migração
def carregar_arquivo():
    # Selecionar o arquivo .rar ou .zip
    file_path = filedialog.askopenfilename(filetypes=[("Arquivos RAR ou ZIP", "*.rar;*.zip")])
    if file_path:
        # Caminho de extração ajustado para a nova pasta B:\migracao\bkp
        extracted_folder_path = r'B:\migracao\bkp'

        try:
            # Descompactar o arquivo
            descompactar_arquivo(file_path, extracted_folder_path)

            # Ler os arquivos CSV de clientes e processos
            clientes_df, processos_df = ler_arquivos_csv(extracted_folder_path)

            # Tratar e unir os dados
            merged_df = tratar_dados(clientes_df, processos_df)

            # Salvar o arquivo migrado no novo diretório B:\migracao\saida
            output_file_path = r'B:\migracao\saida\migracao_advbox.csv'

            # Verificar se o diretório de saída existe, e criar se necessário
            output_dir = r'B:\migracao\saida'
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            arquivo_migrado = salvar_arquivo_migrado(merged_df, output_file_path)

            # Mostrar sucesso
            messagebox.showinfo("Sucesso", f'Arquivo de migração gerado em: {arquivo_migrado}')
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
    
# Interface gráfica
def criar_interface():
    root = Tk()
    root.title("Ferramenta de Migração")
    root.geometry("400x200")

    carregar_button = Button(root, text="Carregar Arquivo de Backup", command=carregar_arquivo)
    carregar_button.pack(pady=20)

    root.mainloop()

# Chamar a interface gráfica
if __name__ == "__main__":
    criar_interface()
