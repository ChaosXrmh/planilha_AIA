import os
import sys
import re
import pandas as pd
from pathlib import Path
import base64


def _normalize_col(name: str) -> str:
    """Normaliza nomes de colunas: minúsculas, sem acentos, sem espaços/pontuação."""
    if not isinstance(name, str):
        return ''
    s = name.lower()
    # remove acentuação básica
    s = s.replace('ã', 'a').replace('á', 'a').replace('à', 'a').replace('â', 'a')
    s = s.replace('é', 'e').replace('è', 'e').replace('ê', 'e')
    s = s.replace('í', 'i').replace('ì', 'i').replace('î', 'i')
    s = s.replace('ó', 'o').replace('ò', 'o').replace('õ', 'o').replace('ô', 'o')
    s = s.replace('ú', 'u').replace('ù', 'u').replace('û', 'u')
    s = s.replace('ç', 'c')
    # keep only alphanumerics
    s = re.sub(r'[^a-z0-9]', '', s)
    return s


def _find_column(df, alternatives):
    """Procura uma coluna no DataFrame a partir de alternativas (lista de nomes possíveis).
    Retorna o nome real da coluna ou None.
    """
    norm_map = {col: _normalize_col(col) for col in df.columns}
    alts_norm = [_normalize_col(a) for a in alternatives]
    for real, norm in norm_map.items():
        if norm in alts_norm:
            return real
    # tenta correspondência por substring
    for alt in alts_norm:
        for real, norm in norm_map.items():
            if alt in norm:
                return real
    return None
# ... (seus imports continuam iguais) ...

# ============================================================================
# CONFIGURAÇÕES INICIAIS (AJUSTADO PARA EXE)
# ============================================================================

# Lógica para descobrir onde o programa está rodando
if getattr(sys, 'frozen', False):
    # Se for um executável (.exe), pegamos o caminho do executável
    SCRIPT_DIR = Path(sys.executable).parent
else:
    # Se for script normal (.py), pegamos o caminho do arquivo
    # Ajuste aqui conforme sua estrutura de pastas original
    # Se o script está na mesma pasta do Excel, use .parent apenas uma vez
    SCRIPT_DIR = Path(__file__).parent.parent 

    # Diretório onde ficam os arquivos Excel a serem processados
    DATA_DIR = SCRIPT_DIR / "data"

    # Valor padrão (pode ser sobrescrito em tempo de execução)
    NOME_ARQUIVO_ORIGINAL = "data/Numeração FALE SEMPRE 081225.xlsx"
    # CAMINHO_ARQUIVO será definido em tempo de execução quando o usuário escolher o arquivo
    CAMINHO_ARQUIVO = SCRIPT_DIR / NOME_ARQUIVO_ORIGINAL
    PASTA_SAIDA = SCRIPT_DIR / "CREFITECH"  # valor padrão — será sobrescrito em tempo de execução para uploads_<empresa>
TAMANHO_LOTE = 100
# ============================================================================
# FUNÇÃO PARA CRIAR PASTA
# ============================================================================

def criar_pasta_saida():
    """Cria a pasta de saída se não existir."""
    try:
        PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
        print(f"✓ Pasta de saída confirmada: {PASTA_SAIDA}")
    except Exception as e:
        print(f"✗ Erro ao criar pasta: {e}")
        sys.exit(1)

# ============================================================================
# FUNÇÃO PARA VALIDAR ARQUIVO
# ============================================================================

def validar_arquivo():
    """Valida se o arquivo Excel existe e é acessível."""
    if not CAMINHO_ARQUIVO.exists():
        print(f"✗ Erro: Arquivo não encontrado em:")
        print(f"  {CAMINHO_ARQUIVO}")
        sys.exit(1)
    
    if not CAMINHO_ARQUIVO.is_file():
        print(f"✗ Erro: {CAMINHO_ARQUIVO} não é um arquivo válido")
        sys.exit(1)
    
    print(f"✓ Arquivo encontrado: {CAMINHO_ARQUIVO.name}")

# ============================================================================
# FUNÇÃO PARA CARREGAR DADOS
# ============================================================================

def carregar_dados():
    """Carrega os dados do Excel com tratamento de erros."""
    try:
        print("\n📂 Lendo arquivo Excel... Aguarde.")
        df = pd.read_excel(CAMINHO_ARQUIVO)

        # Padroniza nomes das colunas para minúsculas e remove espaços
        df.columns = df.columns.str.lower().str.strip()
        
        # Validação do DataFrame
        if df.empty:
            print("✗ Erro: O arquivo Excel está vazio!")
            sys.exit(1)
        
        total_linhas = len(df)
        total_colunas = len(df.columns)
        
        print(f"✓ Arquivo carregado com sucesso!")
        print(f"  └─ Total de linhas: {total_linhas:,}")
        print(f"  └─ Total de colunas: {total_colunas}")
        print(f"  └─ Colunas: {', '.join(df.columns.tolist())}")
        
        return df
    
    except Exception as e:
        print(f"✗ Erro ao carregar arquivo: {e}")
        sys.exit(1)

# ============================================================================
# FUNÇÃO PARA DIVIDIR E SALVAR ARQUIVOS
# ============================================================================

def selecionar_e_formatar_dados(df):
    """Seleciona apenas as 3 colunas necessárias e formata com os tipos corretos."""
    try:
        # caso o Excel tenha importado um CSV inteiro em UMA coluna (ex.: 'numero,acao,cnpj'),
        # dividir essa coluna por delimitador comum e reconstruir o DataFrame
        if df.shape[1] == 1:
            first_val = None
            if len(df) > 0:
                first_val = df.iloc[0, 0]
            # detectar delimitador simples
            delim = None
            for d in [',', ';', '\t']:
                if isinstance(first_val, str) and d in first_val:
                    delim = d
                    break
            if delim:
                splitted = df[df.columns[0]].astype(str).str.split(delim, expand=True)
                # checar se primeira linha é header (contém palavras como 'numero'/'acao'/'cnpj')
                header_row = [s.strip() for s in splitted.iloc[0].tolist()]
                header_norms = [_normalize_col(x) for x in header_row]
                if any(h in ('numero', 'acao', 'cnpj', 'cpfcnpj', 'taxid', 'did') for h in header_norms):
                    # usa primeira linha como header
                    new_df = splitted.copy()
                    new_df.columns = header_row
                    new_df = new_df.drop(index=0).reset_index(drop=True)
                else:
                    # cria nomes genéricos
                    new_df = splitted
                    new_df.columns = [f'col{i+1}' for i in range(new_df.shape[1])]
                df = new_df

        # tenta identificar colunas equivalentes
        numero_col = _find_column(df, ['numero', 'num', 'did', 'id', 'numeroid'])
        cnpj_col = _find_column(df, ['cnpj', 'cpf/cnpj', 'cpfcnpj', 'cpf', 'taxid', 'taxidnumber'])
        acao_col = _find_column(df, ['acao', 'action', 'operacao', 'operacao'])

        if not numero_col or not cnpj_col:
            print(f"✗ Erro: Colunas necessárias não encontradas. Esperadas algo como 'numero' e 'cnpj'.")
            print(f"   Colunas disponíveis: {df.columns.tolist()}")
            # ao invés de sair, retorna erro controlado
            raise ValueError(f"Colunas necessárias faltando. Disponíveis: {df.columns.tolist()}")

        # Seleciona as colunas encontradas e renomeia para os nomes padrão
        df_selected = df.copy()
        cols_to_take = [numero_col, cnpj_col]
        if acao_col:
            cols_to_take.insert(1, acao_col)
        df_selected = df_selected[cols_to_take]
        rename_map = {numero_col: 'numero', cnpj_col: 'cnpj'}
        if acao_col:
            rename_map[acao_col] = 'acao'
        df_selected = df_selected.rename(columns=rename_map)


        # Se não existe coluna 'acao', crie e preencha com valor global (se existir)
        if 'acao' not in df_selected.columns:
            user_action = globals().get('SELECTED_ACTION')
            fill_val = user_action if user_action else ''
            df_selected['acao'] = fill_val

        # Limpeza e normalização do campo 'numero': remover quaisquer caracteres não-dígitos
        # e remover prefixo de país '55' caso exista (manter apenas DDD + número local)
        df_selected['numero'] = df_selected['numero'].astype(str).str.replace(r'\D', '', regex=True)
        df_selected['numero'] = df_selected['numero'].apply(lambda s: s[2:] if s.startswith('55') else s)

        # Formata cada coluna
        df_selected['numero'] = pd.to_numeric(df_selected['numero'], errors='coerce').astype('Int64')
        df_selected['acao'] = df_selected['acao'].astype(str)
        df_selected['cnpj'] = df_selected['cnpj'].astype(str)

        # Garante a ordem correta das colunas de saída
        df_selected = df_selected[['numero', 'acao', 'cnpj']]

        print("✓ Dados formatados (mapeamento automático de colunas):")
        print(f"  └─ NUMERO coluna original: {numero_col}")
        print(f"  └─ CNPJ coluna original: {cnpj_col}")
        if acao_col:
            print(f"  └─ ACAO coluna original: {acao_col}")
        else:
            print(f"  └─ ACAO: criada/definida com: {globals().get('SELECTED_ACTION', '')}")

        return df_selected
    except Exception as e:
        print(f"✗ Erro ao formatar dados: {e}")
        raise

def dividir_e_salvar(df):
    """Divide o DataFrame em lotes e salva em arquivos CSV."""
    total_linhas = len(df)
    contador_arquivo = 1
    arquivos_criados = []
    
    # Seleciona e formata os dados antes de dividir
    df_formatado = selecionar_e_formatar_dados(df)
    # prefixo padronizado recebido via variável global (definida em main)
    file_prefix = globals().get('FILE_PREFIX', 'Cadastro_numeros_CREFITECH')
    
    print(f"\n📝 Dividindo em lotes de {TAMANHO_LOTE} linhas...")
    print(f"   Será criado aproximadamente {(total_linhas // TAMANHO_LOTE) + 1} arquivo(s)\n")
    
    try:
        for i in range(0, total_linhas, TAMANHO_LOTE):
            # Extrai o lote
            fatia = df_formatado.iloc[i : i + TAMANHO_LOTE]
            
            # Define o caminho de saída
            numero_padronizado = str(contador_arquivo).zfill(3)  # Adiciona zeros à esquerda (001, 002...)
            nome_saida = PASTA_SAIDA / f"{file_prefix}_{numero_padronizado}.csv"
            
            # Salva em CSV com separador de vírgula
            fatia.to_csv(nome_saida, index=False, encoding='utf-8-sig', sep=',')
            
            arquivos_criados.append(nome_saida)
            linhas_lote = len(fatia)
            porcentagem = (i + linhas_lote) / total_linhas * 100
            
            print(f"  {contador_arquivo:3d}. {nome_saida.name:50s} ({linhas_lote:3d} linhas) - {porcentagem:5.1f}%")
            
            contador_arquivo += 1
        
        return arquivos_criados
    
    except Exception as e:
        print(f"\n✗ Erro ao salvar arquivos: {e}")
        sys.exit(1)

# ============================================================================
# FUNÇÃO PRINCIPAL
# ============================================================================

def main():
    """Executa o fluxo principal do programa."""
    print("=" * 80)
    print("SISTEMA DE DIVISÃO DE LOTES - CREFITECH")
    print("=" * 80)
    # Pergunta ao usuário qual ação deseja realizar
    action = None
    while action not in ('criar', 'alterar', 'deletar'):
        escolha = input("Escolha ação - (C)riar, (A)lterar, (D)eletar: ").strip().lower()
        if not escolha:
            continue
        chave = escolha[0]
        if chave == 'c':
            action = 'criar'
        elif chave == 'a':
            action = 'alterar'
        elif chave == 'd':
            action = 'deletar'
        else:
            print("Opção inválida. Digite C, A ou D.")

    # Pergunta o nome da empresa
    company = ''
    while not company:
        company_raw = input("Informe o nome da empresa (ex: SURF): ").strip()
        if not company_raw:
            print("Nome da empresa não pode ficar vazio.")
            continue
        # sanitiza nome (remove caracteres inválidos e espaços)
        company = re.sub(r'[^A-Za-z0-9_-]', '', company_raw.replace(' ', '_'))
        if not company:
            print("Nome da empresa contém apenas caracteres inválidos. Tente outro.")

    # Define prefixo conforme a ação escolhida
    prefix_map = {
        'criar': 'Cadastro_numeros',
        'alterar': 'Alterar_numeros',
        'deletar': 'Deletar_numeros'
    }
    prefix = prefix_map.get(action, 'Cadastro_numeros')

    # Global para ser usada na função dividir_e_salvar
    globals()['FILE_PREFIX'] = f"{prefix}_{company}"

    # Ajusta a pasta de saída para uploads_<empresa>
    uploads_folder = SCRIPT_DIR / f"uploads_{company}"
    globals()['PASTA_SAIDA'] = uploads_folder

    # Salva a ação escolhida para ser utilizada na formatação dos dados
    globals()['SELECTED_ACTION'] = action

    # Lista arquivos Excel disponíveis na pasta data e permite seleção
    data_dir = DATA_DIR
    excel_files = []
    if data_dir.exists() and data_dir.is_dir():
        for p in sorted(data_dir.iterdir()):
            if p.is_file() and p.suffix.lower() in ('.xlsx', '.xls'):
                excel_files.append(p)

    if not excel_files:
        print(f"✗ Nenhum arquivo Excel encontrado em: {data_dir}")
        print("Coloque o arquivo na pasta 'data' ou informe o caminho manualmente.")
        # permite que usuário informe caminho completo
        manual = input("Informe o caminho completo do arquivo Excel: ").strip()
        if not manual:
            print("Nenhum arquivo informado. Encerrando.")
            sys.exit(1)
        chosen_path = Path(manual)
    else:
        print("\nArquivos Excel encontrados:")
        for idx, p in enumerate(excel_files, start=1):
            print(f"  {idx}. {p.name}")

        choice = None
        while choice is None:
            sel = input(f"Escolha o arquivo pelo número (1-{len(excel_files)}) ou 'm' para caminho manual: ").strip().lower()
            if sel == 'm':
                manual = input("Informe o caminho completo do arquivo Excel: ").strip()
                if manual:
                    chosen_path = Path(manual)
                    break
                else:
                    continue
            if sel.isdigit():
                n = int(sel)
                if 1 <= n <= len(excel_files):
                    chosen_path = excel_files[n-1]
                    break
            print("Opção inválida.")

    # Define CAMINHO_ARQUIVO global para uso pelas funções
    globals()['CAMINHO_ARQUIVO'] = chosen_path
    print(f"\nArquivo selecionado: {chosen_path}")

    # Executa as etapas
    validar_arquivo()
    criar_pasta_saida()
    df = carregar_dados()
    arquivos_criados = dividir_e_salvar(df)
    
    # Resumo final
    print("\n" + "=" * 80)
    print("✓ PROCESSO FINALIZADO COM SUCESSO!")
    print("=" * 80)
    print(f"\n📊 Resumo:")
    print(f"  └─ Total de arquivo(s) criado(s): {len(arquivos_criados)}")
    print(f"  └─ Total de linhas processadas: {len(df):,}")
    print(f"  └─ Local de saída: {PASTA_SAIDA}")
    print("\n✨ Todos os arquivos estão prontos para importação!\n")

# ============================================================================
# PONTO DE ENTRADA
# ============================================================================

if __name__ == "__main__":
    main()


def processar_arquivo_excel(caminho_arquivo_entrada, acao, empresa_raw, tamanho_lote, pasta_base_saida):
    """
    Função principal adaptada para ser chamada por uma API.
    Recebe todos os parâmetros necessários e retorna um dicionário com o resultado.
    """
    try:
        # Sanitização
        company = re.sub(r'[^A-Za-z0-9_-]', '', empresa_raw.replace(' ', '_'))
        if not company:
            return {"success": False, "error": "Nome da empresa inválido."}

        prefix_map = {
            'criar': 'Cadastro_numeros',
            'alterar': 'Alterar_numeros',
            'deletar': 'Deletar_numeros'
        }
        prefix = prefix_map.get(acao.lower(), 'Cadastro_numeros')
        file_prefix = f"{prefix}_{company}"

        # Pasta de saída
        pasta_saida_final = Path(pasta_base_saida) / f"uploads_{company}"
        pasta_saida_final.mkdir(parents=True, exist_ok=True)

        # Carrega o arquivo Excel
        df = pd.read_excel(caminho_arquivo_entrada)
        # tenta usar a função de seleção/formatacao que faz mapeamento automático
        try:
            df_sel = selecionar_e_formatar_dados(df)
        except Exception as e:
            return {"success": False, "error": f"Erro ao mapear/formatar colunas: {e}"}

        # Sobrescreve a ação conforme parâmetro (garante consistência)
        df_sel['acao'] = acao.lower()

        total_linhas = len(df_sel)
        arquivos_criados = []

        # Ajusta tamanho de lote
        try:
            tamanho_lote = int(tamanho_lote)
        except Exception:
            tamanho_lote = TAMANHO_LOTE
        if tamanho_lote <= 0:
            tamanho_lote = TAMANHO_LOTE

        contador_arquivo = 1
        for i in range(0, total_linhas, tamanho_lote):
            fatia = df_sel.iloc[i: i + tamanho_lote]
            numero_padronizado = str(contador_arquivo).zfill(3)
            nome_saida = pasta_saida_final / f"{file_prefix}_{numero_padronizado}.csv"
            fatia.to_csv(nome_saida, index=False, encoding='utf-8-sig', sep=',')
            arquivos_criados.append(str(nome_saida.name))
            contador_arquivo += 1

        # Empacota conteúdo dos arquivos para enviar ao cliente (base64)
        files_data = []
        for p in arquivos_criados:
            fullpath = pasta_saida_final / p
            try:
                with open(fullpath, 'rb') as fh:
                    data = fh.read()
                b64 = base64.b64encode(data).decode('ascii')
                files_data.append({
                    'name': p,
                    'content_b64': b64
                })
            except Exception:
                # se falhar ao ler, ainda inclui o nome
                files_data.append({'name': p, 'content_b64': None})

        return {
            "success": True,
            "total_files": len(arquivos_criados),
            "total_lines": total_linhas,
            "output_folder": str(pasta_saida_final),
            "files": arquivos_criados,
            "files_data": files_data
        }

    except Exception as e:
        return {"success": False, "error": str(e)}