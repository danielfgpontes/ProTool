import os
import json
import base64
import io
import re
import streamlit as st
import tempfile
from openai import OpenAI
import fitz  # PyMuPDF
import openpyxl
from datetime import datetime
import requests
from pathlib import Path

# --- CONFIGURAÇÃO INICIAL (OBRIGATORIAMENTE O PRIMEIRO COMANDO) ---
st.set_page_config(page_title="Análise Documental - Prefeituras", layout="wide")

# Tratamento seguro caso o arquivo busca_cad.py não esteja no mesmo diretório
try:
    from busca_cad import BuscaTripla
except ImportError:
    st.warning("⚠️ Módulo 'busca_cad' não encontrado no diretório. Usando classe simulada para contingência.")
    class BuscaTripla:
        def iniciar(self, path): return True
        def busca(self, rua, lote, quadra): return {'sucesso': False, 'zona_final': 'ZH2'}

# --- CONFIGURAÇÃO DO MODELO ---
MODEL = "gpt-5.4-mini" 
MODELO_ARQUIVO = "MODELO_Memorial_Planilha_Declaração_R00.xlsx"
CAMINHO_MAPA_DXF = "mapa_zoneamento4.dxf"

@st.cache_resource(show_spinner="Carregando mapa cadastral na memória do servidor...")
def inicializar_motor_busca():
    buscador = BuscaTripla()
    if os.path.exists(CAMINHO_MAPA_DXF):
        buscador.iniciar(CAMINHO_MAPA_DXF)
        return buscador, True
    return buscador, False

motor_busca, mapa_disponivel = inicializar_motor_busca()

# Inicialização do Session State para persistência estável
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
if 'cache_ceps' not in st.session_state:
    st.session_state['cache_ceps'] = {}
if 'tags_debug' not in st.session_state:
    st.session_state['tags_debug'] = {}
if 'zona_final' not in st.session_state:
    st.session_state['zona_final'] = None
if 'sucesso_processamento' not in st.session_state:
    st.session_state['sucesso_processamento'] = False

# --- FUNÇÕES AUXILIARES DE LIMPEZA E FORMATAÇÃO ---

def formatar_numero_br(valor):
    if valor is None: return ""
    valor_str = str(valor).strip()
    apenas_numeros = re.sub(r'[^\d.,]', '', valor_str)
    if not apenas_numeros: return ""
    qtd_pontos = apenas_numeros.count('.')
    qtd_virgulas = apenas_numeros.count(',')

    if qtd_pontos > 0 and qtd_virgulas > 0:
        pos_ponto = apenas_numeros.rfind('.')
        pos_virgula = apenas_numeros.rfind(',')
        if pos_ponto > pos_virgula:
            apenas_numeros = apenas_numeros.replace(',', '').replace('.', ',')
        else:
            apenas_numeros = apenas_numeros.replace('.', '')
    elif qtd_pontos > 0:
        if qtd_pontos > 1 or (len(apenas_numeros) - apenas_numeros.rfind('.') > 3 and len(apenas_numeros) > 4):
            apenas_numeros = apenas_numeros.replace('.', '')
        else:
            apenas_numeros = apenas_numeros.replace('.', ',')
    elif qtd_virgulas > 1:
        partes = apenas_numeros.rsplit(',', 1)
        apenas_numeros = partes[0].replace(',', '') + ',' + partes[1]
    return apenas_numeros

def formatar_documento_detalhado(valor):
    if valor is None or str(valor).strip() == "": return "CPF/CNPJ:", ""
    numeros = re.sub(r'\D', '', str(valor))
    if not numeros: return "CPF/CNPJ:", str(valor)
    if len(numeros) <= 11:
        numeros = numeros.zfill(11)
        return "CPF:", f"{numeros[:3]}.{numeros[3:6]}.{numeros[6:9]}-{numeros[9:]}"
    else:
        numeros = numeros.zfill(14)
        return "CNPJ:", f"{numeros[:2]}.{numeros[2:5]}.{numeros[5:8]}/{numeros[8:12]}-{numeros[12:]}"

def formatar_cep(cep_bruto):
    if not cep_bruto: return ""
    cep_limpo = re.sub(r'\D', '', str(cep_bruto))
    if len(cep_limpo) != 8: return cep_bruto
    return f"{cep_limpo[:5]}-{cep_limpo[5:]}"

def extrair_termos_significativos(texto):
    termos_ignorar = [
        'bairro', 'loteamento', 'jardim', 'jd', 'residencial', 'comercial',
        'industrial', 'parque', 'condomínio', 'cond', 'avenida', 'av', 'rua',
        'r', 'alameda', 'travessa', 'pça', 'praça', 'rodovia', 'estrada',
        'via', 'trav', 'estr', 'rod', 'de', 'do', 'da', 'dos', 'das', 'e', 'ou'
    ]
    if not texto: return []
    texto_limpo = re.sub(r'\d+', '', str(texto)).strip()
    texto_limpo = re.sub(r'[^\w\s]', ' ', texto_limpo)
    texto_limpo = re.sub(r'\s+', ' ', texto_limpo).strip()
    palavras = texto_limpo.upper().split()
    return [p for p in palavras if p.lower() not in termos_ignorar]

def selecionar_cep_por_loteamento(resultados_cep, loteamento):
    if not resultados_cep: return None
    if len(resultados_cep) == 1: return resultados_cep[0]
    termos_lote = set(extrair_termos_significativos(loteamento))
    pontuacoes = []
    for idx, resultado in enumerate(resultados_cep):
        termos_bairro = set(extrair_termos_significativos(resultado.get('bairro', '')))
        coincidencias = termos_lote.intersection(termos_bairro)
        pontuacoes.append({'indice': idx, 'resultado': resultado, 'pontuacao': len(coincidencias)})
    pontuacoes_ordenadas = sorted(pontuacoes, key=lambda x: x['pontuacao'], reverse=True)
    return pontuacoes_ordenadas[0]['resultado']

def limpar_rua_para_cep(rua):
    termos_ignorar = ['av', 'avenida', 'rua', 'r', 'alameda', 'travessa', 'praca', 'praça']
    if not rua: return ""
    rua_limpa = re.sub(r'\d+', '', str(rua)).strip()
    rua_limpa = re.sub(r'[^\w\s]', ' ', rua_limpa)
    rua_limpa = re.sub(r'\s+', ' ', rua_limpa).strip()
    palavras = rua_limpa.upper().split()
    if not palavras: return ""
    palavras_significativas = [p for p in palavras if p.lower() not in termos_ignorar]
    return ' '.join(palavras_significativas) if palavras_significativas else rua_limpa

def buscar_cep_viacep(rua_limpa, cidade_limpa, estado_limpo):
    try:
        url_viacep = f"https://viacep.com.br/ws/{estado_limpo}/{cidade_limpa}/{rua_limpa}/json/"
        response = requests.get(url_viacep, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        if response.status_code == 200:
            dados = response.json()
            if isinstance(dados, (list, dict)) and 'erro' not in dados: return dados
        return None
    except: return None

def buscar_cep_por_endereco(rua, cidade, estado, loteamento=""):
    try:
        if not rua or not cidade or not estado: return ""
        chave_cache = f"{rua}|{cidade}|{estado}".upper()
        if chave_cache in st.session_state['cache_ceps']:
            return st.session_state['cache_ceps'][chave_cache]
        
        rua_limpa = limpar_rua_para_cep(rua)
        cidade_limpa = str(cidade).strip().upper()
        estado_limpo = str(estado).strip().upper()[:2]
        
        dados = buscar_cep_viacep(rua_limpa, cidade_limpa, estado_limpo)
        if dados and isinstance(dados, list):
            resultado_selecionado = selecionar_cep_por_loteamento(dados, loteamento) if len(dados) > 1 and loteamento else dados[0]
            cep = resultado_selecionado.get('cep', '')
            if cep:
                cep_formatado = formatar_cep(cep)
                st.session_state['cache_ceps'][chave_cache] = cep_formatado
                return cep_formatado
        elif dados and isinstance(dados, dict):
            cep = dados.get('cep', '')
            if cep:
                cep_formatado = formatar_cep(cep)
                st.session_state['cache_ceps'][chave_cache] = cep_formatado
                return cep_formatado
        return ""
    except: return ""

def obter_pasta_origem(uploaded_file):
    try:
        home = str(Path.home())
        desktop = Path(home) / "Desktop"
        if desktop.exists(): return str(desktop)
        return home
    except: return None

def salvar_arquivo_xlsx(arquivo_bytes, nome_arquivo, pasta_destino):
    try:
        pasta_path = Path(pasta_destino)
        pasta_path.mkdir(parents=True, exist_ok=True)
        caminho_completo = pasta_path / nome_arquivo
        with open(caminho_completo, 'wb') as f:
            f.write(arquivo_bytes.getvalue())
        return str(caminho_completo)
    except: return None

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
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for i, page in enumerate(doc):
            if i >= 5: break
            pix = page.get_pixmap(dpi=150)
            b64 = base64.b64encode(pix.tobytes("jpeg")).decode('utf-8')
            user_content.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}})
        doc.close()

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
            doc = fitz.open(stream=f_bytes, filetype="pdf")
            for i, page in enumerate(doc):
                if i >= 5: break
                pix = page.get_pixmap(dpi=150)
                b64 = base64.b64encode(pix.tobytes("jpeg")).decode('utf-8')
                user_content.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}})
            doc.close()
    
    response = client.chat.completions.create(
        model=MODEL,
        messages=[{"role": "system", "content": "Retorne estritamente JSON válido."}, {"role": "user", "content": user_content}],
        response_format={"type": "json_object"},
        temperature=0.1
    )
    return json.loads(response.choices[0].message.content)

def preencher_aba_com_tags(ws, contexto):
    tags_encontradas, tags_substituidas = {}, {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                for k, v in contexto.items():
                    tag = "{?" + k + "}" if "{?" + k + "}" in cell.value else "{{" + k + "}}"
                    if tag in cell.value:
                        texto_substituto = str(v) if v is not None else ""
                        if tag not in tags_encontradas: tags_encontradas[tag] = []
                        tags_encontradas[tag].append({'celula': cell.coordinate, 'valor': texto_substituto})
                        cell.value = cell.value.replace(tag, texto_substituto)
                        tags_substituidas[tag] = tags_substituidas.get(tag, 0) + 1
    return tags_encontradas, tags_substituidas

def criar_contexto_dados(dados, zona_dxf=None):
    mat = dados.get('matricula', {})
    idf1 = dados.get('identificacao_1', {})
    idf2 = dados.get('identificacao_2', {})
    prj = dados.get('projeto', {})
    
    tipo_doc_1, num_doc_1 = formatar_documento_detalhado(idf1.get('cnpj_cpf'))
    tipo_doc_2, num_doc_2 = formatar_documento_detalhado(idf2.get('cnpj_cpf')) if idf2.get('cnpj_cpf') else ("", "")
    
    agora = datetime.now()
    meses = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    
    # SOLUÇÃO CEP PARTE 1: Limpa a string da rua antes de enviar para o motor de busca e API de CEP
    rua_bruta = str(mat.get('confrontacao_frente', '')).strip()
    rua = re.split(r'(?i),|\smedindo', rua_bruta)[0].strip()
    
    cidade = str(mat.get('cidade', '')).strip().upper()
    estado = str(mat.get('estado', '')).strip().upper()
    loteamento = str(mat.get('loteamento', '')).strip()
    
    cep = buscar_cep_por_endereco(rua, cidade, estado, loteamento) if rua and cidade and estado else ""
    
    contexto = {
        'cnm': mat.get('cnm', ''), 'matricula': mat.get('matricula', ''), 'folha': mat.get('folha', ''),
        'cartorio': mat.get('cartorio', ''), 'livro': mat.get('livro', ''), 'data_registro': mat.get('data_registro', ''),
        'data_documento': mat.get('data_documento', ''), 'loteamento': loteamento.upper(), 'lote': str(mat.get('lote', '')).upper(),
        'quadra': str(mat.get('quadra', '')).upper(), 'cidade': cidade, 'estado': estado,
        'area': formatar_numero_br(mat.get('area', '')), 'area_construida_total': formatar_numero_br(prj.get('area_construida_total', '')),
        'numero_pavimentos': formatar_numero_br(prj.get('numero_pavimentos', '')), 'numero_vagas': formatar_numero_br(prj.get('numero_vagas', '')),
        'inclinacao_telhado': formatar_numero_br(prj.get('inclinacao_telhado', '')), 'altura_maxima': formatar_numero_br(prj.get('altura_maxima', '')),
        'confrontacao_frente': rua.upper(), 'confrontacao_fundos': str(mat.get('confrontacao_fundos', '')).upper(),
        'confrontacao_lado_direito': str(mat.get('confrontacao_lado_direito', '')).upper(), 'confrontacao_lado_esquerdo': str(mat.get('confrontacao_lado_esquerdo', '')).upper(),
        'proprietario_1': str(idf1.get('proprietario', '')).upper(), 'tipo_doc_1': tipo_doc_1.replace(':', ''), 'num_doc_1': num_doc_1, 'doc_completo_1': f"{tipo_doc_1} {num_doc_1}" if num_doc_1 else "",
        'proprietario_2': str(idf2.get('proprietario', '')).upper() if idf2.get('proprietario') else "", 'tipo_doc_2': tipo_doc_2.replace(':', ''), 'num_doc_2': num_doc_2, 'doc_completo_2': f"{tipo_doc_2} {num_doc_2}" if num_doc_2 else "",
        'finalidade_obra': str(prj.get('finalidade_obra', '')).upper(), 'desenhista': str(prj.get('desenhista', '')).upper(), 'tipo_telhado': str(prj.get('tipo_telhado', '')).upper(),
        'tipo_forro': str(prj.get('tipo_forro', '')).upper(), 'endereco_obra': str(prj.get('endereco_obra', '')).upper(), 'cep': cep, 'zona_zoneamento': zona_dxf if zona_dxf else "ZH2",
        'data_atual': agora.strftime("%d/%m/%Y"), 'data_extensa': f"Sorriso - MT, {agora.day:02d} de {meses[agora.month]} de {agora.year}"
    }
    return contexto

def gerar_arquivo_final(dados, zona_dxf=None):
    contexto = criar_contexto_dados(dados, zona_dxf)
    try:
        wb_completo = openpyxl.load_workbook(MODELO_ARQUIVO)
        todas_as_tags = {}
        for nome_aba in wb_completo.sheetnames:
            ws = wb_completo[nome_aba]
            tags_encontradas, tags_substituidas = preencher_aba_com_tags(ws, contexto)
            if tags_encontradas or tags_substituidas:
                todas_as_tags[nome_aba] = {'encontradas': tags_encontradas, 'substituidas': tags_substituidas}
        
        xlsx_preenchido_io = io.BytesIO()
        wb_completo.save(xlsx_preenchido_io)
        xlsx_preenchido_io.seek(0)
        wb_completo.close()
        st.session_state['tags_debug'] = todas_as_tags
        return {'arquivo_xlsx': xlsx_preenchido_io, 'nome_proprietario': contexto['proprietario_1']}
    except Exception as e:
        raise Exception(f"Erro ao gerar arquivo final: {str(e)}")

# --- INTERFACE STREAMLIT ---
st.title("🏛️ Automação Documental - Prefeituras")
st.markdown("*Selecione os arquivos necessários para iniciar a extração.*")

api_key = st.text_input("OpenAI API Key:", type="password")
if not api_key: api_key = os.environ.get("OPENAI_API_KEY")

col1, col2, col3 = st.columns(3)
with col1: f_mat = st.file_uploader("📋 Matrícula do Imóvel", type=["pdf", "png", "jpg", "jpeg"], key="matricula")
with col2: f_idf_1 = st.file_uploader("🪪 Identificação - Proprietário 1", type=["pdf", "png", "jpg", "jpeg"], key="idf_1")
with col3: f_idf_2 = st.file_uploader("🪪 Identificação - Proprietário 2 (Opcional)", type=["pdf", "png", "jpg", "jpeg"], key="idf_2")

f_prj = st.file_uploader("📐 Pranchas Arquitetônicas", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="projeto")

if st.button("⚙️ Processar e Gerar Documentos", type="primary", use_container_width=True):
    if not api_key:
        st.error("❌ Por favor, insira a chave da OpenAI para continuar.")
        st.stop()
    if not f_mat or not f_idf_1 or not f_prj:
        st.error("❌ Por favor, preencha todos os campos obrigatórios.")
        st.stop()
    
    # Limpa estados de sucesso anteriores para evitar falsos positivos
    st.session_state['sucesso_processamento'] = False
    
    pasta_origem = obter_pasta_origem(f_mat)
    st.session_state['pasta_origem'] = pasta_origem
        
    prompts = {
        "matricula": "Analise a matrícula do imóvel e extraia os seguintes dados em JSON: cnm, matricula, folha, cartorio, livro, data_registro, data_documento, loteamento, lote, quadra, cidade, estado, area (apenas o número com vírgula), confrontacao_frente, confrontacao_fundos, confrontacao_lado_direito, confrontacao_lado_esquerdo.",
        "identificacao": "Analise o documento de identificação e extraia: proprietario (nome completo), cnpj_cpf.",
        "projeto": "Analise as pranchas arquitetônicas e extraia em JSON: area_construida_total, numero_pavimentos, numero_vagas, finalidade_obra, desenhista, tipo_telhado, inclinacao_telhado, tipo_forro, altura_maxima, endereco_obra."
    }

    with st.spinner("⏳ Processando inteligência artificial..."):
        try:
            res_mat = extract_data_from_document(prompts["matricula"], f_mat, api_key)
            res_idf_1 = extract_data_from_document(prompts["identificacao"], f_idf_1, api_key)
            res_idf_2 = extract_data_from_document(prompts["identificacao"], f_idf_2, api_key) if f_idf_2 else {}
            res_prj = extract_data_from_multiple_files(prompts["projeto"], f_prj, api_key)
            
            st.session_state['dados_extraidos'] = {"matricula": res_mat, "identificacao_1": res_idf_1, "identificacao_2": res_idf_2, "projeto": res_prj}
            
            # Cálculo de Zoneamento via DXF
            if mapa_disponivel:
                rua_bruta = str(res_mat.get('confrontacao_frente', '')).strip()
                rua_alvo = re.split(r'(?i),|\smedindo', rua_bruta)[0].strip()
                res_zona = motor_busca.busca(rua_alvo, res_mat.get('lote'), res_mat.get('quadra'))
                st.session_state['zona_final'] = res_zona['zona_final'] if res_zona.get('sucesso') else "ZH2"
            else:
                st.session_state['zona_final'] = "ZH2"
            
            # Geração da Planilha
            resultado = gerar_arquivo_final(st.session_state['dados_extraidos'], st.session_state['zona_final'])
            st.session_state['arquivo_gerado'] = resultado['arquivo_xlsx']
            st.session_state['nome_proprietario'] = resultado['nome_proprietario'] if resultado['nome_proprietario'] else "DOC"
            
            if pasta_origem:
                nome_arquivo = f"{st.session_state['nome_proprietario']}_Memorial_Planilha_Declaração_R00.xlsx"
                st.session_state['caminho_arquivo_salvo'] = salvar_arquivo_xlsx(st.session_state['arquivo_gerado'], nome_arquivo, pasta_origem)
            
            # ATIVAÇÃO DO SINALIZADOR DE SUCESSO PERSISTENTE
            st.session_state['sucesso_processamento'] = True
            st.rerun()
            
        except Exception as e:
            st.error(f"❌ Erro crítico: {e}")

# --- BLOCOS DE EXIBIÇÃO PERSISTENTES (FORA DO BOTÃO) ---
if st.session_state['sucesso_processamento']:
    st.divider()
    st.subheader("📄 Resultados da Análise Documental")
    
    # Exibe informações sem sumir com interações na página
    st.success(f"📍 Lote localizado com sucesso! Zona de Zoneamento definida: **{st.session_state['zona_final']}**")
    
    if st.session_state['caminho_arquivo_salvo']:
        st.success(f"💾 Arquivo salvo automaticamente no servidor em:\n`{st.session_state['caminho_arquivo_salvo']}`")
    else:
        st.warning("⚠️ Não foi possível salvar automaticamente no Desktop local do servidor. Faça o download manual abaixo.")
        
    st.download_button(
        label="📥 Baixar Planilha Excel Preenchida (.XLSX)",
        data=st.session_state['arquivo_gerado'],
        file_name=f"{st.session_state['nome_proprietario']}_Memorial_Planilha_Declaração_R00.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="xlsx_download",
        use_container_width=True
    )

    with st.expander("📋 Ver Dados Brutos Extraídos (JSON)", expanded=False):
        st.json(st.session_state['dados_extraidos'])
