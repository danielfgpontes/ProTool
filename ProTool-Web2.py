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
from copy import copy

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
    if not cep_bruto:
        return ""
    cep_limpo = re.sub(r'\D', '', str(cep_bruto))
    if len(cep_limpo) != 8:
        return cep_bruto
    return f"{cep_limpo[:5]}-{cep_limpo[5:]}"

def extrair_termos_significativos(texto):
    termos_ignorar = [
        'bairro', 'loteamento', 'jardim', 'jd', 'residencial', 'comercial',
        'industrial', 'parque', 'condomínio', 'cond', 'avenida', 'av', 'rua',
        'r', 'alameda', 'travessa', 'pça', 'praça', 'rodovia', 'estrada',
        'via', 'trav', 'estr', 'rod', 'de', 'do', 'da', 'dos', 'das',
        'e', 'ou', 'norte', 'sul', 'leste', 'oeste', 'n', 's', 'l', 'o'
    ]
    if not texto:
        return []
    texto_limpo = re.sub(r'\d+', '', str(texto)).strip()
    texto_limpo = re.sub(r'[^\w\s]', ' ', texto_limpo)
    texto_limpo = re.sub(r'\s+', ' ', texto_limpo).strip()
    palavras = texto_limpo.upper().split()
    return [p for p in palavras if p.lower() not in termos_ignorar]

def selecionar_cep_por_loteamento(resultados_cep, loteamento):
    if not resultados_cep:
        return None
    if len(resultados_cep) == 1:
        return resultados_cep[0]
    
    termos_lote = set(extrair_termos_significativos(loteamento))
    st.info(f"🔍 Comparando {len(resultados_cep)} resultados de CEP com loteamento: **{loteamento}**")
    
    pontuacoes = []
    for idx, resultado in enumerate(resultados_cep):
        termos_bairro = set(extrair_termos_significativos(resultado.get('bairro', '')))
        coincidencias = termos_lote.intersection(termos_bairro)
        pontuacao = len(coincidencias)
        pontuacoes.append({
            'indice': idx,
            'resultado': resultado,
            'pontuacao': pontuacao
        })
    
    pontuacoes_ordenadas = sorted(pontuacoes, key=lambda x: x['pontuacao'], reverse=True)
    
    if pontuacoes_ordenadas[0]['pontuacao'] > 0:
        res = pontuacoes_ordenadas[0]['resultado']
        st.success(f"✅ CEP selecionado baseado no bairro: `{formatar_cep(res.get('cep'))}`")
        return res
    return resultados_cep[0]

def limpar_rua_para_cep(rua):
    termos_ignorar = [
        'norte', 'sul', 'leste', 'oeste', 'n', 's', 'l', 'o',
        'av', 'avenida', 'rua', 'r', 'alameda', 'travessa', 'praca', 'praça',
        'rod', 'rodovia', 'estr', 'estrada', 'via', 'pça', 'trav'
    ]
    if not rua:
        return ""
    rua_limpa = re.sub(r'\d+', '', str(rua)).strip()
    rua_limpa = re.sub(r'[^\w\s]', ' ', rua_limpa)
    rua_limpa = re.sub(r'\s+', ' ', rua_limpa).strip()
    palavras = rua_limpa.upper().split()
    
    if not palavras:
        return ""
    
    palavras_significativas = []
    for palavra in reversed(palavras):
        if palavra.lower() not in termos_ignorar:
            palavras_significativas.append(palavra)
        elif palavras_significativas:
            break
            
    palavras_significativas.reverse()
    resultado = ' '.join(palavras_significativas)
    return resultado if resultado else rua_limpa

def buscar_cep_por_endereco(rua, cidade, estado, loteamento=""):
    try:
        if not rua or not cidade or not estado:
            return ""
        
        rua_limpa = limpar_rua_para_cep(rua)
        cidade_limpa = str(cidade).strip().upper()
        estado_limpo = str(estado).strip().upper()[:2]
        
        if not rua_limpa or len(rua_limpa) < 3:
            return ""
        
        url = f"https://viacep.com.br/ws/{estado_limpo}/{cidade_limpa}/{rua_limpa}/json/"
        response = requests.get(url, timeout=5)
        
        if response.status_code == 200:
            dados = response.json()
            if isinstance(dados, dict) and 'erro' in dados:
                return ""
            if isinstance(dados, list) and len(dados) > 0:
                if len(dados) > 1 and loteamento:
                    resultado_selecionado = selecionar_cep_por_loteamento(dados, loteamento)
                else:
                    resultado_selecionado = dados[0]
                return formatar_cep(resultado_selecionado.get('cep', ''))
            elif isinstance(dados, dict):
                return formatar_cep(dados.get('cep', ''))
        return ""
    except:
        return ""

def obter_pasta_origem(uploaded_file):
    try:
        if hasattr(uploaded_file, 'name'):
            home = str(Path.home())
            desktop = Path(home) / "Desktop"
            if desktop.exists():
                return str(desktop)
            return home
    except:
        pass
    return None

def salvar_arquivo_xlsx(arquivo_bytes, nome_arquivo, pasta_destino):
    try:
        pasta_path = Path(pasta_destino)
        pasta_path.mkdir(parents=True, exist_ok=True)
        caminho_completo = pasta_path / nome_arquivo
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
            raise ValueError(f"Erro ao processar o PDF: {e}")

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
    Preenche a aba localizando as tags {{chave}} e substituindo o valor,
    garantindo que os estilos e formatos originais da célula permaneçam intactos.
    """
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                texto_original = cell.value
                texto_modificado = texto_original
                houve_alteracao = False
                
                for k, v in contexto.items():
                    tag = "{?" + k + "}" if "{?" in texto_modificado else "{{" + k + "}}"
                    if tag in texto_modificado:
                        texto_substituto = str(v) if v is not None else ""
                        texto_modificado = texto_modificado.replace(tag, texto_substituto)
                        houve_alteracao = True
                
                if houve_alteracao:
                    # No openpyxl, alterar apenas o atributo .value PRESERVA os estilos 
                    # existentes na célula (fontes, cores, bordas, formatos numéricos).
                    # Não precisamos clonar os objetos com a biblioteca 'copy'.
                    cell.value = texto_modificado

def criar_contexto_dados(dados):
    mat = dados.get('matricula', {})
    idf1 = dados.get('identificacao_1', {})
    idf2 = dados.get('identificacao_2', {})
    prj = dados.get('projeto', {})
    
    tipo_doc_1, num_doc_1 = formatar_documento_detalhado(idf1.get('cnpj_cpf'))
    tipo_doc_2, num_doc_2 = formatar_documento_detalhado(idf2.get('cnpj_cpf')) if idf2.get('cnpj_cpf') else ("", "")
    
    agora = datetime.now()
    meses = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    
    rua = str(mat.get('confrontacao_frente', '')).strip()
    cidade = str(mat.get('cidade', '')).strip().upper()
    estado = str(mat.get('estado', '')).strip().upper()
    loteamento = str(mat.get('loteamento', '')).strip()
    
    cep = buscar_cep_por_endereco(rua, city := cidade, estado, loteamento) if rua and cidade and estado else ""
    
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
        'confrontacao_fundos': str(mat.get('confrontacao_fundos', '')).upper(),
        'confrontacao_lado_direito': str(mat.get('confrontacao_lado_direito', '')).upper(),
        'confrontacao_lado_esquerdo': str(mat.get('confrontacao_lado_esquerdo', '')).upper(),

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
        'cep': cep,
        
        'data_atual': agora.strftime("%d/%m/%Y"),
        'data_extensa': f"Sorriso - MT, {agora.day:02d} de {meses[agora.month]} de {agora.year}"
    }
    return contexto

def gerar_arquivo_final(dados):
    contexto = criar_contexto_dados(dados)
    proprietario_1 = contexto['proprietario_1']
    
    try:
        # Forçamos a abertura preservando todos os estilos e estruturas complexas do arquivo XML
        wb_completo = openpyxl.load_workbook(MODELO_ARQUIVO, data_only=False, keep_vba=True)
        
        for nome_aba in wb_completo.sheetnames:
            ws = wb_completo[nome_aba]
            preencher_aba_com_tags(ws, contexto)
        
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
st.success("✅ O arquivo gerado manterá estritamente o layout do seu modelo original.")

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
    
    if st.session_state['caminho_arquivo_salvo']:
        st.success(f"✅ Arquivo salvo em: `{st.session_state['caminho_arquivo_salvo']}`")
    else:
        st.info("💾 Você pode fazer download do arquivo abaixo:")
    
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

    st.divider()
    with st.expander("Ver Dados Brutos Extraídos (JSON)"):
        st.json(st.session_state['dados_extraidos'])
