import glob
import inspect
import json
import openpyxl
import os
import pandas as pd
import re
import requests
import sys
import time
from colorama import Fore, Style
from datetime import datetime, timedelta

# Obtém os argumentos da linha de comando, excluindo o nome do script
argumentos = sys.argv[1:]

# Defina esta variável como True para habilitar a exibição dos prints
SHOW_RESPONSE = True

# Define o header a ser usado de forma global no script
HEADER_AGENT = "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0"

# Diretório base do arquivo atual
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)  # Mudando o diretório de trabalho atual para o diretório base

# Diretório temporário base para arquivos
BASE_TEMP = os.path.join(BASE_DIR, ".tmp_rel")
# Diretório base para arquivos
BASE_FILES = os.path.join(BASE_DIR, "files")

# Cria Diretorios se não existir
os.makedirs(BASE_FILES, exist_ok=True)

# Inicializa listas para armazenar contratos e relatórios
contratos = []
relatorios = []
csv_files = []
data_referencia = None


def process_csv():
    # Localiza os arquivos terminados com "_lst_inf.csv" e coloca em uma lista.
    csv_files = glob.glob(os.path.join(BASE_FILES, "*_lst_inf.csv"))

    def remove_first_lines(file_path, num_lines):
        with open(file_path, "r", encoding="latin1") as f:
            lines = f.readlines()

        # Adicionar o ponto e vírgula na sétima linha
        if len(lines) >= 7:
            lines[6] = lines[6].rstrip() + ";\n"

        with open(file_path, "w", encoding="latin1") as f:
            f.writelines(lines[num_lines:])

    # Itera sobre a lista de Arquivos CSV
    for arquivo_csv in csv_files:
        print(f"{head_log()} Lendo arquivo: {arquivo_csv}")

        # Condição especifica
        if "dnitms" in arquivo_csv:
            # remover as três primeiras linhas do arquivo CSV
            remove_first_lines(arquivo_csv, 6)

        # leia o arquivo CSV
        df = pd.read_csv(arquivo_csv, delimiter=";", encoding="latin1", header=0)

        # verificar se o nome do arquivo contém "dnitms"
        if "dnitms" in arquivo_csv:
            # remover a 20ª coluna do DataFrame
            df.drop(df.columns[19], axis=1, inplace=True)

        # obter o nome do arquivo sem extensão
        nome_arquivo_sem_extensao = os.path.splitext(os.path.basename(arquivo_csv))[0]

        # substitua os caracteres inválidos do nome da planilha por "_"
        sheet_name = re.sub("[^A-Za-z0-9]+", "_", nome_arquivo_sem_extensao)[:30]

        # crie um arquivo XLSX e adicione os dados do DataFrame a ele
        xls_name = os.path.join(BASE_FILES, f"{nome_arquivo_sem_extensao}.xlsx")
        print(f"{head_log()} Escrevendo no arquivo: {xls_name}")
        with pd.ExcelWriter(xls_name, engine="openpyxl") as writer:
            ws = writer.book.create_sheet(sheet_name, 0)
            df.to_excel(
                writer, sheet_name=sheet_name, index=False, header=True, startrow=0
            )

        # verifique se o número de linhas e colunas do arquivo CSV é igual ao número de linhas
        # e colunas do arquivo XLSX
        df_xlsx = pd.read_excel(xls_name, sheet_name=sheet_name, header=0)
        num_linhas_csv, num_colunas_csv = df.shape
        num_linhas_xlsx, num_colunas_xlsx = df_xlsx.shape

        print(f"{head_log()} Validando arquivo: {xls_name}")
        if num_linhas_csv == num_linhas_xlsx and num_colunas_csv == num_colunas_xlsx:
            # exclua o arquivo CSV
            print(f"{head_log()} Arquivo {xls_name} validado.")
            print(f"{head_log()} Excluindo arquivo: {arquivo_csv}")
            os.remove(arquivo_csv)


def head_log():
    # Obtém o nome do método atual a partir da pilha de chamadas
    frame = inspect.currentframe().f_back  # type: ignore
    actual_method = frame.f_code.co_name  # type: ignore
    report_head = f"{Fore.LIGHTWHITE_EX}{datetime.now()} | {Fore.GREEN}{Style.BRIGHT}{actual_method} | {Style.RESET_ALL}"
    return report_head


def load_contracts():
    contracts_file = os.path.join(BASE_TEMP, "contracts.json")
    if not os.path.isfile(contracts_file):
        if SHOW_RESPONSE:
            print(
                f"{head_log()} {Fore.RED}Arquivo contracts.json não encontrado em {BASE_TEMP}{Fore.RESET}"
            )
        return None
    with open(contracts_file, "r") as file:
        return json.load(file)


def permited_reports(report):
    contracts_data = load_contracts()
    if contracts_data is None:
        if SHOW_RESPONSE:
            print(
                f"{head_log()} {Fore.RED}Arquivo contracts.json não encontrado.{Fore.RESET}"
            )
        return False

    permited_reports = contracts_data.get("permited_reports", [])
    for report_type in permited_reports:
        if report_type == report:
            return True

    if SHOW_RESPONSE:
        print(
            f'{head_log()} {Fore.RED}Relatório "{report}" não configurado ou não existe.{Fore.RESET}'
        )
    return False


def read_contract(contract_name):
    contracts_data = load_contracts()
    if contracts_data is None:
        if SHOW_RESPONSE:
            print(
                f"{head_log()} {Fore.RED}Arquivo contracts.json não encontrado.{Fore.RESET}"
            )
        return None

    for contract in contracts_data.get("contracts", []):
        if contract["contract"] == contract_name:
            return contract["base_url"], contract["types_report"]

    if SHOW_RESPONSE:
        print(
            f'{head_log()} {Fore.RED}Contrato "{contract_name}" não encontrado em contracts.json.{Fore.RESET}'
        )
    return False


def extrair_contratos():
    contracts_data = (
        load_contracts()
    )  # Supondo que load_contracts() seja uma função válida que retorna os dados dos contratos
    contratos = []
    # Verifica se a chave "contracts" está presente no JSON
    if "contracts" in contracts_data:  # type: ignore
        # Verifica se "contracts" é uma lista de dicionários
        if isinstance(contracts_data["contracts"], list):  # type: ignore
            # Itera sobre cada contrato na lista de contratos
            for contrato in contracts_data["contracts"]:  # type: ignore
                # Verifica se o contrato é um dicionário e se contém a chave "contract"
                if isinstance(contrato, dict) and "contract" in contrato:
                    # Adiciona o valor do campo "contract" à lista de contratos
                    contratos.append(contrato["contract"])
    return contratos


def show_data(user, base_url, contract, start_date, end_date, base_filename_full):
    if SHOW_RESPONSE:
        print(f"{head_log()} User: {user}, Base URL: {base_url}, contract: {contract}")
        print(
            f"{head_log()} Start Date: {start_date}, End Date: {end_date}, Base File: {base_filename_full}"
        )


def mount_base_filename_full(base_filename, contract, count_file, report):
    return f"{base_filename}_{contract}_{count_file}_{report}"


def read_credentials(contract_name):
    contract_info = read_contract(contract_name)
    if contract_info is None:
        return None

    credentials_file = os.path.join(BASE_TEMP, f"credentials_{contract_name}.json")
    if not os.path.isfile(credentials_file):
        if SHOW_RESPONSE:
            print(
                f"{head_log()} {Fore.RED}Arquivo credentials_{contract_name}.json não encontrado{Fore.RESET}"
            )
        with open(credentials_file, "w") as file:
            json.dump({"user": None, "password": None}, file, indent=4)
        return None

    with open(credentials_file, "r") as file:
        credentials_data = json.load(file)
        user = credentials_data.get("user")
        password = credentials_data.get("password")
        if user is not None and password is not None:
            return user, password

    if SHOW_RESPONSE:
        print(
            f"{head_log()} {Fore.RED}Conteúdo de usuário e senha no arquivo credentials_{contract_name}.json é nulo.{Fore.RESET}"
        )
    return None


def get_date(pr_year=None, pr_month=None):
    if pr_year and pr_month:
        start_date_base = datetime(pr_year, pr_month, 1)
        start_date_filename = f'{start_date_base.strftime("%Y%m")}'
        start_date = start_date_base.strftime("%d/%m/%Y")
        last_day_of_month = (
            datetime(pr_year, pr_month, 1) + timedelta(days=32)
        ).replace(day=1) - timedelta(days=1)
        end_date = last_day_of_month.strftime("%d/%m/%Y")
    else:
        yesterday = datetime.now() - timedelta(days=1)
        start_date_base = datetime(yesterday.year, yesterday.month, 1)
        start_date_filename = f'{yesterday.strftime("%Y%m")}'
        start_date = start_date_base.strftime("%d/%m/%Y")
        end_date = yesterday.strftime("%d/%m/%Y")
    today_formatted = datetime.now().strftime("%Y%m%d")
    base_filename = (
        f"{start_date_filename}_{today_formatted}_{datetime.now().strftime('%H%M%S')}"
    )
    return start_date, end_date, base_filename


def check_connection(url_contract):
    try:
        headers = {"User-Agent": HEADER_AGENT}
        response = requests.get(url_contract, headers=headers, verify=False)
        return response.status_code == 200
    except requests.exceptions.RequestException as e:
        if SHOW_RESPONSE:
            print(
                f"{head_log()} {Fore.RED}Erro de requisição - Verifique a VPN: {e}{Fore.RESET}"
            )
        return False


def login(base_url, username, password, session):
    # Faz a requisição POST para autenticação
    data = {"login": username, "senha": password, "btn_entrar": ""}
    headers = {"User-Agent": HEADER_AGENT}
    response = session.post(f"{base_url}/Autenticar", data=data, verify=False)
    if response.status_code == 200:
        if SHOW_RESPONSE:
            print(f"{head_log()} {response.status_code} | Session: {session}")
        if "Usuário e/ou senha inválidos!" in response.text:
            if SHOW_RESPONSE:
                print(
                    f"{head_log()} {Fore.RED}A senha no arquivo de credenciais está errada.{Fore.RESET}"
                )
            return False
        else:
            return True
    else:
        if SHOW_RESPONSE:
            print(f"{head_log()} {Fore.RED}Falha na solicitação de login.{Fore.RESET}")
        return False


def is_excel_file(output_file):
    try:
        if os.path.isfile(output_file):
            if SHOW_RESPONSE:
                print(f"{head_log()} Caminho do Arquivo: {output_file}")
            with open(output_file, "rb") as f:
                magic_number = f.read(8)
                if SHOW_RESPONSE:
                    print(f"{head_log()} Magic Number: {magic_number}")
                return magic_number in [
                    b"\x09\x08\x10\x00\x00\x06\x05\x00",  # xls (CFB)
                    b"\x50\x4B\x03\x04\x14\x00\x06\x00",  # xlsx (primeiros 16 bytes do arquivo zip)
                    b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1",  # CFB (outro)
                ]
    except FileNotFoundError:
        if SHOW_RESPONSE:
            print(
                f'{head_log()} {Fore.RED}Erro: O sistema não pode encontrar o arquivo "{output_file}".{Fore.RESET}'
            )
        return False


def healt_report(
    session,
    base_url,
    base_filename,
    start_date,
    end_date,
    somente_imagens,
    tipo_origem_dado=0,
):
    output_file = os.path.join(BASE_FILES, f"{base_filename}.xls")
    if SHOW_RESPONSE:
        print(f"{head_log()} {output_file}")

    url_mounted = f"{base_url}/relatorio/RelatorioFuncionamento"
    headers = {"User-Agent": HEADER_AGENT}
    data = {
        "tipo_relatorio": "",
        "dataini": start_date,
        "datafim": end_date,
        "grupo_equipamento": 0,
        "tipo_origem_dado": tipo_origem_dado,
        "somente_imagens": somente_imagens,
        "utiliza_movel": 0,
        "tipo_arquivo": "XLS",
    }
    requisitions_report(session, output_file, url_mounted, headers, data)


def rel_tst(session, base_url, base_filename, start_date, end_date):
    print(f"{head_log()} {base_filename}")
    healt_report(session, base_url, base_filename, start_date, end_date, 1)


def rel_inf(session, base_url, base_filename, start_date, end_date):
    print(f"{head_log()} {base_filename}")
    healt_report(session, base_url, base_filename, start_date, end_date, 2)


def rel_flx(session, base_url, base_filename, start_date, end_date):
    print(f"{head_log()} {base_filename}")
    healt_report(session, base_url, base_filename, start_date, end_date, 0, 1)


def lst_inf(session, base_url, base_filename, start_date, end_date, contract=None):
    headers = {"User-Agent": HEADER_AGENT}
    if contract == "dnitms":
        output_file = os.path.join(BASE_FILES, f"{base_filename}.csv")
        url_mounted = f"{base_url}/relatorio/RelatorioDinamicoServlet"
        if SHOW_RESPONSE:
            print(f"{head_log()} {output_file}")
        data = {
            "limpar_tela": "",
            "relatorio_dinamico": 11,
            "data_ini": start_date,
            "horaini": "00:00",
            "data_fim": end_date,
            "horafim": "23:59",
            "com_cabecalho": 1,
        }
    else:
        output_file = os.path.join(BASE_FILES, f"{base_filename}.xls")
        if SHOW_RESPONSE:
            print(f"{head_log()} {output_file}")
        url_mounted = f"{base_url}/ferramenta/ListarInfracao"
        data = {
            "id_infracao": "",
            "auto": "",
            "id_imagem": "",
            "data_infracao_ini": start_date,
            "hora_infracao_ini": "00:00",
            "data_infracao_fim": end_date,
            "hora_infracao_fim": "23:59",
            "numero_lote": "",
            "placa": "",
            "velocidade": "",
            "sel_cod_pista_cliente": "",
            "cod_pista": "",
            "id_grupo_equipamento": 0,
            "id_processo": 0,
            "id_inconsistencia": "",
            "id_enquadramento": 0,
            "id_grupo_autuador": "",
            "id_classe": "",
            "id_usuario_atual": "",
            "id_processo_historico": 0,
            "id_usuario_historico": "",
            "data_processamento_ini": "",
            "hora_processamento_ini": "",
            "data_processamento_fim": "",
            "hora_processamento_fim": "",
            "id_inconsistencia_historico": "",
            "classif_inc_processamento": "",
            "filtrar_resultados": 0,
            "max_linhas": 100,
            "ignorar_imagem_teste": True,
            "acao": "exportar_csv",
            "para_processo": "",
            "valor_customizado": 2,
        }
    requisitions_report(session, output_file, url_mounted, headers, data)


def rel_lap(session, base_url, base_filename, start_date, end_date, grupo_sub_periodo):
    output_file = os.path.join(BASE_FILES, f"{base_filename}.xls")
    if SHOW_RESPONSE:
        print(f"{head_log()} {output_file}")

    url_mounted = f"{base_url}/relatorio/RelatorioLAP"
    headers = {"User-Agent": HEADER_AGENT}
    data = {
        "dataini": start_date,
        "datafim": end_date,
        "grupo_sub_periodo": grupo_sub_periodo,
        "tipo_ordenacao": 0,
        "tipo_arquivo": "XLS",
    }
    requisitions_report(session, output_file, url_mounted, headers, data)


def rel_lapd(session, base_url, base_filename, start_date, end_date):
    print(f"{head_log()} {base_filename}")
    rel_lap(session, base_url, base_filename, start_date, end_date, 3)


def rel_lapn(session, base_url, base_filename, start_date, end_date):
    print(f"{head_log()} {base_filename}")
    rel_lap(session, base_url, base_filename, start_date, end_date, 4)


def rel_lapi(session, base_url, base_filename, start_date, end_date):
    print(f"{head_log()} {base_filename}")
    rel_lap(session, base_url, base_filename, start_date, end_date, 5)


def requisitions_report(session, output_file, url_mounted, headers, data):
    retryes = 18
    timedout_sleep = 20
    loop_count = 0
    while loop_count <= retryes:
        if loop_count == (retryes + 1):
            print(
                f"{head_log()} Loop: {loop_count} {Fore.RED}Erro ao Processar - Tempo decorrido: {(retryes * timedout_sleep)} {Fore.RESET}"
            )
            break
        response = session.post(url_mounted, headers=headers, data=data, verify=False)
        content_disposition = response.headers.get("Content-Disposition", "")
        if response.status_code == 200 and "filename=" in content_disposition:
            print(f"{head_log()} Loop: {loop_count} | Relatorio Localizado")
            with open(output_file, "wb") as file:
                file.write(response.content)
            print(f"{head_log()} Loop: {loop_count} | Relatorio Salvo: ")
            return True
        else:
            print(
                f"{head_log()} Loop: {loop_count} | Tempo decorrido: {(loop_count * timedout_sleep)}s - Aguardando..."
            )
            loop_count += 1
            time.sleep(timedout_sleep)
    return False


def base_loop_relatorios(contracts, reports_imported, year_month=None):
    print(f"{head_log()} Start App...")

    year_data = None
    month_data = None
    if year_month:
        year_data = int(year_month[:-2])
        month_data = int(year_month[4:])

    if not contracts:
        contracts = extrair_contratos()

    for contract in contracts:
        users_data = read_credentials(contract)
        contracts_data = read_contract(contract)

        if users_data and contracts_data:
            # Cria uma sessão
            headers = {"User-Agent": HEADER_AGENT}
            session = requests.Session()
            session.headers.update(headers)

            base_url, _ = contracts_data
            if not check_connection(base_url):
                print(f"{head_log()} {Fore.RED}Verifique sua conexão e VPN{Fore.RESET}")

            print(f"{head_log()} Contrato Selecionado: {contract}")

            user, password = users_data
            if login(base_url, user, password, session):
                start_date, end_date, base_filename = get_date(year_data, month_data)
                if reports_imported:
                    reports = reports_imported
                else:
                    _, reports = contracts_data
                count_file = 0
                for report in reports:
                    count_file += 1
                    if report == "rel_tst":
                        base_filename_full = mount_base_filename_full(
                            base_filename, contract, count_file, report
                        )
                        rel_tst(
                            session, base_url, base_filename_full, start_date, end_date
                        )
                        show_data(
                            user,
                            base_url,
                            contract,
                            start_date,
                            end_date,
                            base_filename_full,
                        )
                    elif report == "rel_inf":
                        base_filename_full = mount_base_filename_full(
                            base_filename, contract, count_file, report
                        )
                        rel_inf(
                            session, base_url, base_filename_full, start_date, end_date
                        )
                        show_data(
                            user,
                            base_url,
                            contract,
                            start_date,
                            end_date,
                            base_filename_full,
                        )
                    elif report == "rel_flx":
                        base_filename_full = mount_base_filename_full(
                            base_filename, contract, count_file, report
                        )
                        rel_flx(
                            session, base_url, base_filename_full, start_date, end_date
                        )
                        show_data(
                            user,
                            base_url,
                            contract,
                            start_date,
                            end_date,
                            base_filename_full,
                        )
                    elif report == "rel_lapd":
                        base_filename_full = mount_base_filename_full(
                            base_filename, contract, count_file, report
                        )
                        rel_lapd(
                            session, base_url, base_filename_full, start_date, end_date
                        )
                        show_data(
                            user,
                            base_url,
                            contract,
                            start_date,
                            end_date,
                            base_filename_full,
                        )
                    elif report == "rel_lapn":
                        base_filename_full = mount_base_filename_full(
                            base_filename, contract, count_file, report
                        )
                        rel_lapn(
                            session, base_url, base_filename_full, start_date, end_date
                        )
                        show_data(
                            user,
                            base_url,
                            contract,
                            start_date,
                            end_date,
                            base_filename_full,
                        )
                    elif report == "rel_lapi":
                        base_filename_full = mount_base_filename_full(
                            base_filename, contract, count_file, report
                        )
                        rel_lapi(
                            session, base_url, base_filename_full, start_date, end_date
                        )
                        show_data(
                            user,
                            base_url,
                            contract,
                            start_date,
                            end_date,
                            base_filename_full,
                        )
                    elif report == "lst_inf":
                        base_filename_full = mount_base_filename_full(
                            base_filename, contract, count_file, report
                        )
                        lst_inf(
                            session,
                            base_url,
                            base_filename_full,
                            start_date,
                            end_date,
                            contract,
                        )
                        show_data(
                            user,
                            base_url,
                            contract,
                            start_date,
                            end_date,
                            base_filename_full,
                        )
            session.close()
            process_csv()
            print(f"{head_log()} Fim Contrato {contract}")
        elif users_data:
            user, password = users_data
            print(
                f"{head_log()} User: {user}, Contract: {contract}, No contract data found"
            )
        else:
            print(f"{head_log()} No user data found for contract: {contract}")


# Itera sobre os argumentos passados
i = 0
while i < len(argumentos):
    arg = argumentos[i]
    if arg in ["-c", "--contrato", "--contratos"]:
        i += 1
        # Verifica se há argumentos subsequentes e os adiciona como contratos
        if i < len(argumentos):
            contratos.extend(argumentos[i].split(","))
    elif arg in ["-r", "--relatorio", "--relatorios"]:
        i += 1
        # Verifica se há argumentos subsequentes e os adiciona como relatórios
        if i < len(argumentos):
            relatorios.extend(argumentos[i].split(","))
    elif arg in ["-d", "--datareferencia", "--datareferencias"]:
        i += 1
        # Verifica se há argumentos subsequentes e define a data de referência
        if i < len(argumentos):
            data_referencia = argumentos[i]
    elif arg in ["--help", "-h"]:
        # Exibe a mensagem de ajuda
        print(
            """
        Descrição:
        O script refactoring.py é usado para processar contratos, relatórios e uma data de referência fornecidos como argumentos da linha de comando.

        Uso:
        python refactoring.py [OPÇÕES]

        Opções:
        -c CONTRATO, --contrato=CONTRATO, --contratos=CONTRATO: Especifica o contrato ou uma lista de contratos separados por vírgula a serem processados.
        -r RELATÓRIO, --relatorio=RELATÓRIO, --relatorios=RELATÓRIO: Especifica o relatório ou uma lista de relatórios separados por vírgula a serem processados.
        -d DATA, --datareferencia=DATA, --datareferencias=DATA: Especifica a data de referência para processamento.

        Exemplos:
        - Para processar os contratos dnitms e msvia, os relatórios rel_tst e rel_inf, e a data de referência 202404:
        python refactoring.py -c dnitms,msvia -r rel_tst,rel_inf -d 202404
        - Para processar apenas o contrato dnitms:
        python refactoring.py --contrato=dnitms
        - Para processar o contrato msvia, os relatórios rel_inf e rel_flx, e a data de referência 202405:
        python refactoring.py -c msvia --relatorio=rel_inf,rel_flx -d 202405
        - Para obter ajuda sobre como usar o script:
        python refactoring.py --help
        """
        )
        # Sai do script após exibir a mensagem de ajuda
        sys.exit(0)
    i += 1

base_loop_relatorios(contratos, relatorios, data_referencia)
