import os
import json
import base64
import io
import re
import streamlit as st
from openai import OpenAI
import fitz  # PyMuPDF
import openpyxl
from datetime import datetime
import requests
from pathlib import Path

# --- CONFIGURAÇÃO INICIAL E MODELO ---
MODEL = "gpt-5.4-mini"
MODELO_ARQUIVO = "MODELO_Memorial_Planilha_Declaração_R00.xlsx"

st.set_page_config(page_title="Análise Documental - Prefeituras", layout="wide")

# Inicialização do Session State para persistência
if 'dados_extraidos' not in st.session_state:
    st.session_state['dados_extraidos'] = None
if 'arquivo_gerado' not in st.session_state:
    st.session_state['arquivo_gerado'] = None
if 'nome_proprietario' not in st.session_state:
    st.session_state['nome_proprietario'] = "DOC"
if 'pasta_origem' not in st.session_state:
    st.session_state['pasta_origem'] = None
if 'caminho_arquivo_salvo' not in st.session_state:
    st.session_state['caminho_arquivo_salvo'] = None

# --- FUNÇÕES AUXILIARES ---
def extrair_apenas_numeros_area(valor):
    """Filtra uma string para retornar apenas os números, vírgulas e pontos."""
    if not valor: 
        return ""
    match = re.search(r'[\d.,]+', str(valor))
    return match.group(0) if match else str(valor)

def formatar_documento_detalhado(valor):
    """Limpa o valor, aplica a máscara e retorna uma tupla: (Tipo do Documento, Número Formatado)"""
    if valor is None or str(valor).strip() == "":
        return "CPF/CNPJ:", ""
    numeros = re.sub(r'\D', '', str(valor))
    if not numeros:
        return "CPF/CNPJ:", str(valor)
    if len(numeros) <= 11:
        numeros = numeros.zfill(11)
        return "CPF:", f"{numeros[:3]}.{numeros[3:6]}.{numeros[6:9]}-{numeros[9:]}"
    else:
        numeros = numeros.zfill(14)
        return "CNPJ:", f"{numeros[:2]}.{numeros[2:5]}.{numeros[5:8]}/{numeros[8:12]}-{numeros[12:]}"

def formatar_cep(cep_bruto):
    """
    Formata o CEP corretamente removendo caracteres especiais primeiro.
    
    Exemplos:
    - "78894588" -> "78894-588"
    - "78894-588" -> "78894-588"
    - "78894--588" -> "78894-588"
    """
    if not cep_bruto:
        return ""
    
    # Remove todos os caracteres que não são dígitos
    cep_limpo = re.sub(r'\D', '', str(cep_bruto))
    
    # Verifica se tem 8 dígitos
    if len(cep_limpo) != 8:
        return cep_bruto  # Retorna como estava se não tiver 8 dígitos
    
    # Formata: XXXXX-XXX
    cep_formatado = f"{cep_limpo[:5]}-{cep_limpo[5:]}"
    
    return cep_formatado

def extrair_termos_significativos(texto):
    """
    Extrai termos significativos de um texto, ignorando palavras genéricas.
    
    Exemplo:
    - "Bairro Parque dos Poderes" -> ["PARQUE", "PODERES"]
    - "Loteamento Residencial Jardim Sul" -> ["RESIDENCIAL", "JARDIM", "SUL"]
    """
    
    termos_ignorar = [
        'bairro', 'loteamento', 'jardim', 'jd', 'residencial', 'comercial',
        'industrial', 'parque', 'condomínio', 'cond', 'avenida', 'av', 'rua',
        'r', 'alameda', 'travessa', 'pça', 'praça', 'rodovia', 'estrada',
        'via', 'trav', 'estr', 'rod', 'de', 'do', 'da', 'dos', 'das',
        'e', 'ou', 'norte', 'sul', 'leste', 'oeste', 'n', 's', 'l', 'o'
    ]
    
    if not texto:
        return []
    
    # Remove números e caracteres especiais
    texto_limpo = re.sub(r'\d+', '', str(texto)).strip()
    texto_limpo = re.sub(r'[^\w\s]', ' ', texto_limpo)
    texto_limpo = re.sub(r'\s+', ' ', texto_limpo).strip()
    
    # Separa em palavras
    palavras = texto_limpo.upper().split()
    
    # Filtra termos significativos
    termos_significativos = [p for p in palavras if p.lower() not in termos_ignorar]
    
    return termos_significativos

def selecionar_cep_por_loteamento(resultados_cep, loteamento):
    """
    Compara múltiplos resultados de CEP com o loteamento e seleciona o mais apropriado.
    
    Args:
        resultados_cep: Lista de resultados da API ViaCEP
        loteamento: Nome do loteamento
    
    Returns:
        Dicionário com o CEP selecionado ou o primeiro resultado se nenhuma correspondência
    """
    
    if not resultados_cep:
        return None
    
    if len(resultados_cep) == 1:
        return resultados_cep[0]
    
    # Extrai termos significativos do loteamento
    termos_lote = set(extrair_termos_significativos(loteamento))
    
    st.info(f"🔍 Comparando {len(resultados_cep)} resultados de CEP com loteamento: **{loteamento}**")
    
    with st.expander("📊 Detalhes da Comparação de Loteamentos", expanded=False):
        st.write(f"**Termos significativos do loteamento:** {termos_lote if termos_lote else 'Nenhum'}")
        
        # Mostra cada resultado e sua pontuação
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.write("**Resultado**")
            for idx, resultado in enumerate(resultados_cep):
                st.write(f"`Opção {idx + 1}`")
        
        with col2:
            st.write("**Bairro**")
            for resultado in resultados_cep:
                st.write(f"{resultado.get('bairro', 'N/A')}")
        
        with col3:
            st.write("**Coincidências**")
            for resultado in resultados_cep:
                termos_bairro = set(extrair_termos_significativos(resultado.get('bairro', '')))
                coincidencias = termos_lote.intersection(termos_bairro)
                st.write(f"{len(coincidencias)} termos")
    
    # Calcula pontuação para cada resultado baseada em coincidências
    pontuacoes = []
    
    for idx, resultado in enumerate(resultados_cep):
        termos_bairro = set(extrair_termos_significativos(resultado.get('bairro', '')))
        
        # Conta coincidências entre loteamento e bairro
        coincidencias = termos_lote.intersection(termos_bairro)
        pontuacao = len(coincidencias)
        
        pontuacoes.append({
            'indice': idx,
            'resultado': resultado,
            'pontuacao': pontuacao,
            'coincidencias': coincidencias
        })
    
    # Ordena por pontuação (maior primeiro)
    pontuacoes_ordenadas = sorted(pontuacoes, key=lambda x: x['pontuacao'], reverse=True)
    
    # Se encontrou correspondência, usa o primeiro
    if pontuacoes_ordenadas[0]['pontuacao'] > 0:
        resultado_selecionado = pontuacoes_ordenadas[0]
        cep_formatado = formatar_cep(resultado_selecionado['resultado'].get('cep'))
        st.success(
            f"✅ CEP selecionado (Opção {resultado_selecionado['indice'] + 1}): "
            f"`{cep_formatado}` - "
            f"**{resultado_selecionado['resultado'].get('bairro')}** "
            f"({resultado_selecionado['pontuacao']} coincidências)"
        )
        return resultado_selecionado['resultado']
    else:
        # Se não encontrou correspondência, usa o primeiro resultado
        cep_formatado = formatar_cep(resultados_cep[0].get('cep'))
        st.warning(
            f"⚠️ Nenhuma correspondência encontrada. "
            f"Usando primeiro resultado: `{cep_formatado}` - "
            f"**{resultados_cep[0].get('bairro')}**"
        )
        return resultados_cep[0]

def limpar_rua_para_cep(rua):
    """
    Limpa o nome da rua para busca de CEP.
    Extrai apenas o ÚLTIMO termo significativo, ignorando direções como Norte, Sul, Leste, Oeste.
    
    Exemplos:
    - "AV. DOS IMIGRANTES SUL" -> "IMIGRANTES"
    - "RUA DAS FLORES NORTE" -> "FLORES"
    - "AVENIDA PAULISTA LESTE" -> "PAULISTA"
    - "RUA JOSE DE ALENCAR OESTE" -> "ALENCAR"
    """
    
    # Lista de termos a ignorar (direções e preposições)
    termos_ignorar = [
        'norte', 'sul', 'leste', 'oeste',
        'n', 's', 'l', 'o',
        'av', 'avenida', 'rua', 'r', 'alameda', 'travessa', 'praca', 'praça',
        'rod', 'rodovia', 'estr', 'estrada', 'via', 'pça', 'trav'
    ]
    
    if not rua:
        return ""
    
    # Remove números
    rua_limpa = re.sub(r'\d+', '', str(rua)).strip()
    
    # Remove caracteres especiais, mantendo apenas letras e espaços
    rua_limpa = re.sub(r'[^\w\s]', ' ', rua_limpa)
    
    # Remove espaços múltiplos
    rua_limpa = re.sub(r'\s+', ' ', rua_limpa).strip()
    
    # Separa em palavras
    palavras = rua_limpa.upper().split()
    
    if not palavras:
        return ""
    
    # Remove termos a ignorar do final para o início
    # Mantemos apenas as palavras significativas
    palavras_significativas = []
    
    for palavra in reversed(palavras):
        # Se a palavra não está na lista de ignorar, adicionamos
        if palavra.lower() not in termos_ignorar:
            palavras_significativas.append(palavra)
        # Se já temos uma palavra significativa e encontramos uma para ignorar, paramos
        elif palavras_significativas:
            break
    
    # Inverte para manter ordem original
    palavras_significativas.reverse()
    
    resultado = ' '.join(palavras_significativas)
    
    return resultado if resultado else rua_limpa

def buscar_cep_por_endereco(rua, cidade, estado, loteamento=""):
    """
    Busca o CEP através da API do ViaCEP usando apenas rua, cidade e estado.
    Se houver múltiplos resultados, seleciona baseado no loteamento.
    
    Args:
        rua: Nome da rua/avenida (pode conter direções)
        cidade: Cidade
        estado: Estado (UF - 2 letras)
        loteamento: Nome do loteamento (para comparação)
    
    Returns:
        CEP formatado (XXXXX-XXX) ou vazio se não encontrado
    """
    try:
        if not rua or not cidade or not estado:
            st.warning("⚠️ Dados insuficientes para buscar CEP (rua, cidade ou estado vazio)")
            return ""
        
        # Limpa a rua usando a nova função
        rua_limpa = limpar_rua_para_cep(rua)
        
        cidade_limpa = str(cidade).strip().upper()
        estado_limpo = str(estado).strip().upper()
        
        # Garante que o estado tem apenas 2 letras
        if len(estado_limpo) > 2:
            estado_limpo = estado_limpo[:2]
        
        if not rua_limpa or len(rua_limpa) < 3:
            st.warning(f"⚠️ Nome da rua muito curto após limpeza: '{rua_limpa}'")
            return ""
        
        # Mostra a estrutura da busca
        with st.expander("🔍 Detalhes da Busca de CEP", expanded=False):
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Dados Originais:**")
                st.write(f"- Rua: `{rua}`")
                st.write(f"- Cidade: `{cidade}`")
                st.write(f"- Estado: `{estado}`")
                st.write(f"- Loteamento: `{loteamento}`")
            
            with col2:
                st.write("**Dados Processados:**")
                st.write(f"- Rua (limpa): `{rua_limpa}`")
                st.write(f"- Cidade: `{cidade_limpa}`")
                st.write(f"- Estado: `{estado_limpo}`")
                st.info("ℹ️ Termos de direção (N, S, L, O) foram removidos do nome da rua")
        
        # API ViaCEP - busca por endereço
        # Formato: https://viacep.com.br/ws/[UF]/[cidade]/[logradouro]/json/
        url = f"https://viacep.com.br/ws/{estado_limpo}/{cidade_limpa}/{rua_limpa}/json/"
        
        st.info(f"🌐 URL da API: `{url}`")
        
        response = requests.get(url, timeout=5)
        
        st.info(f"📡 Status HTTP: {response.status_code}")
        
        if response.status_code == 200:
            dados = response.json()
            
            with st.expander("📋 Resposta da API (JSON)", expanded=False):
                st.json(dados)
            
            # Verifica se retornou um erro
            if isinstance(dados, dict) and 'erro' in dados:
                st.error(f"❌ CEP não encontrado para: {rua_limpa}, {cidade_limpa} - {estado_limpo}")
                st.warning("A API retornou um erro. Verifique se a rua, cidade e estado estão corretos.")
                return ""
            
            # Se é uma lista (múltiplos resultados)
            if isinstance(dados, list) and len(dados) > 0:
                st.success(f"✅ {len(dados)} resultado(s) encontrado(s)")
                
                # Se há múltiplos resultados, seleciona baseado no loteamento
                if len(dados) > 1 and loteamento:
                    resultado_selecionado = selecionar_cep_por_loteamento(dados, loteamento)
                else:
                    resultado_selecionado = dados[0]
                
                cep = resultado_selecionado.get('cep', '')
                if cep:
                    # Formata CEP corretamente
                    cep_formatado = formatar_cep(cep)
                    st.success(f"✅ CEP selecionado: {cep_formatado}")
                    st.info(f"📍 Endereço: {resultado_selecionado.get('logradouro', '')}, {resultado_selecionado.get('bairro', '')}")
                    return cep_formatado
            
            # Se for um dicionário único
            elif isinstance(dados, dict):
                cep = dados.get('cep', '')
                if cep:
                    cep_formatado = formatar_cep(cep)
                    st.success(f"✅ CEP encontrado: {cep_formatado}")
                    st.info(f"📍 Endereço: {dados.get('logradouro', '')}, {dados.get('bairro', '')}")
                    return cep_formatado
        
        st.error(f"❌ Erro ao buscar CEP (Status HTTP: {response.status_code})")
        return ""
    
    except requests.exceptions.Timeout:
        st.error(f"⏱️ Timeout ao buscar CEP (aguardou mais de 5 segundos)")
        return ""
    except requests.exceptions.RequestException as e:
        st.error(f"⚠️ Erro de conexão ao buscar CEP: {str(e)}")
        return ""
    except Exception as e:
        st.error(f"⚠️ Erro inesperado ao buscar CEP: {str(e)}")
        return ""

def obter_pasta_origem(uploaded_file):
    """
    Tenta obter a pasta de origem do arquivo carregado.
    """
    try:
        # Streamlit não fornece acesso direto ao caminho do arquivo
        # Tentamos obter através do atributo name
        if hasattr(uploaded_file, 'name'):
            # O arquivo tem um nome, tentamos usar Desktop como padrão
            home = str(Path.home())
            desktop = Path(home) / "Desktop"
            
            if desktop.exists():
                return str(desktop)
            
            # Se Desktop não existe, usa a pasta do usuário
            return home
    except:
        pass
    
    # Retorna None se não conseguir determinar
    return None

def salvar_arquivo_xlsx(arquivo_bytes, nome_arquivo, pasta_destino):
    """
    Salva o arquivo XLSX na pasta especificada.
    
    Args:
        arquivo_bytes: BytesIO com os dados do arquivo
        nome_arquivo: Nome do arquivo a salvar
        pasta_destino: Caminho da pasta destino
    
    Returns:
        Caminho completo do arquivo salvo ou None se houver erro
    """
    try:
        pasta_path = Path(pasta_destino)
        
        # Cria a pasta se não existir
        pasta_path.mkdir(parents=True, exist_ok=True)
        
        # Caminho completo do arquivo
        caminho_completo = pasta_path / nome_arquivo
        
        # Salva o arquivo
        with open(caminho_completo, 'wb') as f:
            f.write(arquivo_bytes.getvalue())
        
        return str(caminho_completo)
    
    except Exception as e:
        st.error(f"Erro ao salvar arquivo: {e}")
        return None

def extract_data_from_document(prompt_text, uploaded_file, api_key):
    client = OpenAI(api_key=api_key)
    user_content = [{"type": "text", "text": prompt_text}]
    
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)
    file_ext = uploaded_file.name.split('.')[-1].lower()

    if file_ext in ['png', 'jpg', 'jpeg']:
        b64 = base64.b64encode(file_bytes).decode('utf-8')
        user_content.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}})
    elif file_ext == 'pdf':
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            for i, page in enumerate(doc):
                if i >= 5: break
                pix = page.get_pixmap(dpi=150)
                b64 = base64.b64encode(pix.tobytes("jpeg")).decode('utf-8')
                user_content.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}})
            doc.close()
        except Exception as e:
            raise ValueError(f"Erro ao processar o PDF da Matrícula/Identificação: {e}")

    response = client.chat.completions.create(
        model=MODEL,
        messages=[{"role": "system", "content": "Retorne estritamente JSON válido."}, {"role": "user", "content": user_content}],
        response_format={"type": "json_object"},
        temperature=0.1
    )
    return json.loads(response.choices[0].message.content)

def extract_data_from_multiple_files(prompt_text, uploaded_files, api_key):
    client = OpenAI(api_key=api_key)
    user_content = [{"type": "text", "text": prompt_text}]
    
    for f in uploaded_files:
        f_bytes = f.read()
        f.seek(0)
        ext = f.name.split('.')[-1].lower()
        if ext in ['png', 'jpg', 'jpeg']:
            b64 = base64.b64encode(f_bytes).decode('utf-8')
            user_content.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}})
        elif ext == 'pdf':
            try:
                doc = fitz.open(stream=f_bytes, filetype="pdf")
                for i, page in enumerate(doc):
                    if i >= 5: break
                    pix = page.get_pixmap(dpi=150)
                    b64 = base64.b64encode(pix.tobytes("jpeg")).decode('utf-8')
                    user_content.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}})
                doc.close()
            except Exception as e:
                raise ValueError(f"Erro ao processar o PDF do Projeto: {e}")
    
    response = client.chat.completions.create(
        model=MODEL, 
        messages=[{"role": "system", "content": "Retorne estritamente JSON válido."}, {"role": "user", "content": user_content}], 
        response_format={"type": "json_object"},
        temperature=0.1
    )
    return json.loads(response.choices[0].message.content)

def preencher_aba_com_tags(ws, contexto):
    """
    Preenche uma aba inteira procurando por tags {{chave}} e substituindo pelos valores.
    """
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                for k, v in contexto.items():
                    tag = "{{" + k + "}}"
                    if tag in cell.value:
                        texto_substituto = str(v) if v is not None else ""
                        cell.value = cell.value.replace(tag, texto_substituto)

def criar_contexto_dados(dados):
    """
    Cria o dicionário de contexto com todos os dados extraídos.
    Também busca o CEP automaticamente.
    """
    mat = dados.get('matricula', {})
    idf1 = dados.get('identificacao_1', {})
    idf2 = dados.get('identificacao_2', {})
    prj = dados.get('projeto', {})
    
    tipo_doc_1, num_doc_1 = formatar_documento_detalhado(idf1.get('cnpj_cpf'))
    tipo_doc_2, num_doc_2 = formatar_documento_detalhado(idf2.get('cnpj_cpf')) if idf2.get('cnpj_cpf') else ("", "")
    
    # Datas automáticas
    agora = datetime.now()
    meses = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    
    # Busca CEP automaticamente usando apenas rua, cidade e estado
    rua = str(mat.get('confrontacao_frente', '')).strip()  # Usa confrontacao_frente como rua
    cidade = str(mat.get('cidade', '')).strip().upper()
    estado = str(mat.get('estado', '')).strip().upper()
    loteamento = str(mat.get('loteamento', '')).strip()  # Adiciona loteamento para comparação
    
    cep = ""
    if rua and cidade and estado:
        cep = buscar_cep_por_endereco(rua, cidade, estado, loteamento)
    
    contexto = {
        'cnm': mat.get('cnm', ''),
        'matricula': mat.get('matricula', ''),
        'folha': mat.get('folha', ''),
        'cartorio': mat.get('cartorio', ''),
        'livro': mat.get('livro', ''),
        'data_registro': mat.get('data_registro', ''),
        'data_documento': mat.get('data_documento', ''),
        'loteamento': loteamento.upper(),
        'lote': str(mat.get('lote', '')).upper(),
        'quadra': str(mat.get('quadra', '')).upper(),
        'cidade': cidade,
        'estado': estado,
        'area': extrair_apenas_numeros_area(mat.get('area', '')),
        'confrontacao_frente': rua.upper(),
        'confrontacao_fundos': mat.get('confrontacao_fundos', ''),
        'confrontacao_lado_direito': mat.get('confrontacao_lado_direito', ''),
        'confrontacao_lado_esquerdo': mat.get('confrontacao_lado_esquerdo', ''),

        # Proprietário 1
        'proprietario_1': str(idf1.get('proprietario', '')).upper(),
        'tipo_doc_1': tipo_doc_1.replace(':', ''),
        'num_doc_1': num_doc_1,
        'doc_completo_1': f"{tipo_doc_1} {num_doc_1}" if num_doc_1 else "",

        # Proprietário 2
        'proprietario_2': str(idf2.get('proprietario', '')).upper() if idf2.get('proprietario') else "",
        'tipo_doc_2': tipo_doc_2.replace(':', ''),
        'num_doc_2': num_doc_2,
        'doc_completo_2': f"{tipo_doc_2} {num_doc_2}" if num_doc_2 else "",

        'area_construida_total': extrair_apenas_numeros_area(prj.get('area_construida_total', '')),
        'numero_pavimentos': prj.get('numero_pavimentos', ''),
        'numero_vagas': prj.get('numero_vagas', ''),
        'finalidade_obra': str(prj.get('finalidade_obra', '')).upper(),
        'desenhista': str(prj.get('desenhista', '')).upper(),
        'tipo_telhado': str(prj.get('tipo_telhado', '')).upper(),
        'inclinacao_telhado': prj.get('inclinacao_telhado', ''),
        'tipo_forro': str(prj.get('tipo_forro', '')).upper(),
        'altura_maxima': prj.get('altura_maxima', ''),
        'endereco_obra': str(prj.get('endereco_obra', '')).upper(),
        'cep': cep,  # CEP buscado automaticamente
        
        'data_atual': agora.strftime("%d/%m/%Y"),
        'data_extensa': f"Sorriso - MT, {agora.day:02d} de {meses[agora.month]} de {agora.year}"
    }
    
    return contexto

def gerar_arquivo_final(dados):
    """
    1. Carrega o modelo único
    2. Preenche TODAS as abas com as tags
    3. Salva o XLSX completo preenchido
    """
    contexto = criar_contexto_dados(dados)
    proprietario_1 = contexto['proprietario_1']
    
    try:
        # Carrega o modelo
        wb_completo = openpyxl.load_workbook(MODELO_ARQUIVO)
        
        # Preenche TODAS as abas
        for nome_aba in wb_completo.sheetnames:
            ws = wb_completo[nome_aba]
            preencher_aba_com_tags(ws, contexto)
        
        # Salva XLSX preenchido em BytesIO (para download)
        xlsx_preenchido_io = io.BytesIO()
        wb_completo.save(xlsx_preenchido_io)
        xlsx_preenchido_io.seek(0)
        
        wb_completo.close()
        
        return {
            'arquivo_xlsx': xlsx_preenchido_io,
            'nome_proprietario': proprietario_1
        }
    
    except FileNotFoundError:
        raise FileNotFoundError(f"Arquivo '{MODELO_ARQUIVO}' não encontrado")
    except Exception as e:
        raise Exception(f"Erro ao gerar arquivo: {str(e)}")

# --- INTERFACE ---
st.title("Automação Documental - Prefeituras")
st.markdown("*Selecione ou arraste os arquivos para extrair as informações e preencher os templates.*")
st.success("✅ Arquivo será gerado em XLSX e salvo automaticamente na pasta de origem")

api_key = st.text_input("OpenAI API Key (Deixe em branco se configurada no sistema):", type="password")
if not api_key: 
    api_key = os.environ.get("OPENAI_API_KEY")

col1, col2, col3 = st.columns(3)
with col1:
    f_mat = st.file_uploader("Matrícula do Imóvel", type=["pdf", "png", "jpg", "jpeg"], key="matricula")
with col2:
    f_idf_1 = st.file_uploader("Documento de Identificação - Proprietário 1", type=["pdf", "png", "jpg", "jpeg"], key="idf_1")
with col3:
    f_idf_2 = st.file_uploader("Documento de Identificação - Proprietário 2 (Opcional)", type=["pdf", "png", "jpg", "jpeg"], key="idf_2")

col_prj = st.columns(1)[0]
with col_prj:
    f_prj = st.file_uploader("Pranchas Arquitetônicas", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="projeto")

if st.button("Processar e Gerar Documentos", type="primary"):
    if not api_key:
        st.error("Por favor, insira a chave da OpenAI para continuar.")
        st.stop()
    
    if not f_mat or not f_idf_1 or not f_prj:
        st.error("Por favor, preencha os campos obrigatórios: Matrícula, Documento Proprietário 1 e Pranchas Arquitetônicas.")
        st.stop()
    
    # Tenta obter pasta de origem
    pasta_origem = obter_pasta_origem(f_mat)
    st.session_state['pasta_origem'] = pasta_origem
        
    prompts = {
        "matricula": (
            "Analise a matrícula do imóvel (terreno) e extraia os seguintes dados em JSON com exatamente estas chaves: "
            "cnm, matricula, folha, cartorio, livro, data_registro, data_documento, "
            "loteamento, lote, quadra, cidade, estado, area (extraia APENAS o valor numérico, sem a unidade de medida 'm²'), "
            "confrontacao_frente (descrição do limite frontal, ex: 'Av. dos Imigrantes Sul'), "
            "confrontacao_fundos (descrição do limite dos fundos. Atenção: pode haver mais de 1 lote, liste todos, ex: 'Lote 06 e 07'), "
            "confrontacao_lado_direito (descrição do limite direito. Atenção: pode haver mais de 1 lote, liste todos, ex: 'Lotes 19 e 20'), "
            "confrontacao_lado_esquerdo (descrição do limite esquerdo. Atenção: pode haver mais de 1 lote, liste todos, ex: 'Lotes 21 e 22')."
        ),
        "identificacao": (
            "Analise o documento de identificação do proprietário e extraia os seguintes dados em JSON com exatamente estas chaves: "
            "proprietario (nome completo), cnpj_cpf."
        ),
        "projeto": (
            "Analise as pranchas arquitetônicas e extraia os seguintes dados em JSON com exatamente estas chaves: "
            "area_construida_total (extraia APENAS o valor numérico, sem a unidade de medida 'm²'), numero_pavimentos, numero_vagas (vagas de estacionamento), "
            "finalidade_obra, desenhista, tipo_telhado, inclinacao_telhado, tipo_forro, "
            "altura_maxima (altura máxima em metros. Atenção: utilize como referência os desenhos de 'Corte' nas pranchas, considerando como altura máxima a cota vertical de maior valor), "
            "endereco_obra."
        ),
    }

    with st.spinner("Analisando documentos da prefeitura e preenchendo arquivo... Isso pode levar alguns segundos."):
        try:
            res_mat = extract_data_from_document(prompts["matricula"], f_mat, api_key) if f_mat else {}
            res_idf_1 = extract_data_from_document(prompts["identificacao"], f_idf_1, api_key) if f_idf_1 else {}
            res_idf_2 = extract_data_from_document(prompts["identificacao"], f_idf_2, api_key) if f_idf_2 else {}
            res_prj = extract_data_from_multiple_files(prompts["projeto"], f_prj, api_key) if f_prj else {}
            
            st.session_state['dados_extraidos'] = {
                "matricula": res_mat, 
                "identificacao_1": res_idf_1, 
                "identificacao_2": res_idf_2,
                "projeto": res_prj
            }
            
            resultado = gerar_arquivo_final(st.session_state['dados_extraidos'])
            
            st.session_state['arquivo_gerado'] = resultado['arquivo_xlsx']
            st.session_state['nome_proprietario'] = resultado['nome_proprietario'] if resultado['nome_proprietario'] else "DOC"
            
            # Tenta salvar o arquivo na pasta de origem
            if pasta_origem:
                nome_arquivo = f"{st.session_state['nome_proprietario']}_Memorial_Planilha_Declaração_R00.xlsx"
                caminho_salvo = salvar_arquivo_xlsx(st.session_state['arquivo_gerado'], nome_arquivo, pasta_origem)
                
                if caminho_salvo:
                    st.session_state['caminho_arquivo_salvo'] = caminho_salvo
                    st.success(f"✅ Arquivo salvo automaticamente em:\n`{caminho_salvo}`")
                else:
                    st.warning("⚠️ Não foi possível salvar o arquivo automaticamente. Use o botão de download.")
            else:
                st.warning("⚠️ Não foi possível determinar a pasta de origem. Use o botão de download.")
            
            st.success("✅ Análise documental finalizada!")
        except Exception as e:
            st.error(f"Erro no processamento: {e}")
            import traceback
            st.error(traceback.format_exc())

# --- EXIBIÇÃO E BOTÕES DE DOWNLOAD ---
if st.session_state['dados_extraidos']:
    st.divider()
    st.subheader("📄 Documento Pronto para Download")
    
    nome_proprietario = st.session_state['nome_proprietario']
    
    # Mostra status do salvamento automático
    if st.session_state['caminho_arquivo_salvo']:
        st.success(f"✅ Arquivo salvo em: `{st.session_state['caminho_arquivo_salvo']}`")
    else:
        st.info("💾 Você pode fazer download do arquivo abaixo:")
    
    # --- ARQUIVO XLSX COMPLETO ---
    if st.session_state['arquivo_gerado']:
        st.download_button(
            label="📥 Baixar XLSX Preenchido",
            data=st.session_state['arquivo_gerado'],
            file_name=f"{nome_proprietario}_Memorial_Planilha_Declaração_R00.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="xlsx_download"
        )
    else:
        st.error("❌ Erro ao gerar XLSX")

    # --- DADOS EXTRAÍDOS ---
    st.divider()
    with st.expander("Ver Dados Brutos Extraídos (JSON)"):
        st.json(st.session_state['dados_extraidos'])