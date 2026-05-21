import re
import math
import os
import streamlit as st
from typing import List, Tuple, Dict, Optional

try:
    import ezdxf
    from ezdxf import bbox as dxfo_bbox
    HAS_EZDXF = True
except ImportError:
    HAS_EZDXF = False
    st.error("⚠️ Biblioteca 'ezdxf' não encontrada. Execute: `pip install ezdxf`")

try:
    from shapely.geometry import Polygon, Point
    HAS_SHAPELY = True
except ImportError:
    HAS_SHAPELY = False
    st.error("⚠️ Biblioteca 'shapely' não encontrada. Execute: `pip install shapely`")


class BuscaTripla:
    """
    Motor de busca para extrair textos (Rua, Lote, Quadra) em desenhos CAD (DXF)
    e determinar a zona de zoneamento baseada na boundary (limite) da hachura.
    Adaptado para integração com Streamlit.
    """
    
    def __init__(self):
        self.lista1 = []
        self.lista2 = []
        self.lista3 = []
        self.melhor1 = None
        self.melhor2 = None
        self.melhor3 = None
        self.melhorEnt = None
        self.zonaFinal = "ZH2"
        
        # Dados em memória
        self.textos_cache = []
        self.hachuras_cache = []
        self.hachuras_filtradas = []
        self.hachuras_com_boundary = []
        self.arquivo_carregado = None
        
    def normalizar_texto(self, texto: str) -> str:
        """Normaliza o texto removendo acentos e caracteres especiais."""
        if not texto:
            return ""
        
        texto = texto.strip().upper()
        
        mapeamento = {
            'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U',
            'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
            'Ã': 'A', 'ã': 'a', 'Õ': 'O', 'õ': 'o', 'Ç': 'C', 'ç': 'c',
            '-': '', '_': '', '/': '', ',': '', ':': '', ';': ''
        }
        
        for char, substituir in mapeamento.items():
            texto = texto.replace(char, substituir)
        
        return texto
    
    def wcmatch(self, texto: str, padrao: str) -> bool:
        """Simula a função wcmatch do AutoLISP para Python usando Regex."""
        if not texto or not padrao:
            return False
        
        padrao_regex = padrao.replace('.', r'\.')
        padrao_regex = padrao_regex.replace('*', '.*')
        padrao_regex = padrao_regex.replace('?', '.')
        padrao_regex = f"^{padrao_regex}$"
        
        return bool(re.match(padrao_regex, texto))
    
    def distance(self, p1: Tuple[float, float], p2: Tuple[float, float]) -> float:
        """Calcula a distância euclidiana entre dois pontos."""
        return math.sqrt((p1[0] - p2[0])**2 + (p1[1] - p2[1])**2)
    
    def carregar_textos_de_dxf(self, caminho_dxf: str) -> List[dict]:
        textos = []
        if not HAS_EZDXF or not os.path.exists(caminho_dxf):
            return textos
        
        try:
            dwg = ezdxf.readfile(caminho_dxf)
            mspace = dwg.modelspace()
            
            for entity in mspace:
                try:
                    if entity.dxftype() in ('TEXT', 'MTEXT', 'ATTRIB'):
                        texto = entity.dxf.text
                        x = entity.dxf.insert[0]
                        y = entity.dxf.insert[1]
                        
                        if texto.strip():
                            textos.append({
                                'texto': texto,
                                'texto_limpo': self.normalizar_texto(texto),
                                'x': float(x),
                                'y': float(y),
                                'nome': f"{entity.dxftype()}_{len(textos)}",
                                'tipo': entity.dxftype(),
                                'camada': entity.dxf.layer
                            })
                except Exception:
                    continue
            
            st.write(f"📄 **{len(textos)}** entidades de texto indexadas na memória.")
            return textos
            
        except Exception as e:
            st.error(f"❌ Erro ao ler entidades de texto do DXF: {e}")
            return []

    def extrair_poligono_hachura(self, entity) -> Optional[Polygon]:
        """Tenta extrair o contorno REAL da hachura ao invés de usar apenas o retângulo da BBox."""
        try:
            if not hasattr(entity, 'paths') or not entity.paths:
                return None
            
            # O primeiro path geralmente é a borda externa da hachura
            path_externo = entity.paths[0]
            pontos = []
            
            if hasattr(path_externo, 'vertices'):
                for v in path_externo.vertices:
                    pontos.append((float(v[0]), float(v[1])))
            elif hasattr(path_externo, 'edges'):
                for edge in path_externo.edges:
                    if hasattr(edge, 'start'):
                        pontos.append((float(edge.start[0]), float(edge.start[1])))
            
            if len(pontos) >= 3:
                poly = Polygon(pontos)
                # Corrige auto-interseções no polígono caso a hachura cad seja irregular
                if not poly.is_valid:
                    poly = poly.buffer(0)
                return poly
        except Exception:
            pass
        return None

    def obter_coordenadas_hachura(self, entity) -> Tuple[float, float, dict]:
        x, y = 0.0, 0.0
        bbox_info = {'xmin': 0, 'xmax': 0, 'ymin': 0, 'ymax': 0}
        
        if not HAS_EZDXF:
            return x, y, bbox_info

        try:
            extents = dxfo_bbox.extents([entity])
            if extents.has_data:
                xmin, ymin = extents.extmin.x, extents.extmin.y
                xmax, ymax = extents.extmax.x, extents.extmax.y
                
                x = (xmin + xmax) / 2.0
                y = (ymin + ymax) / 2.0
                bbox_info = {'xmin': float(xmin), 'xmax': float(xmax), 'ymin': float(ymin), 'ymax': float(ymax)}
                return float(x), float(y), bbox_info
        except:
            pass
        
        return float(x), float(y), bbox_info
    
    def carregar_hachuras_de_dxf(self, caminho_dxf: str) -> List[dict]:
        hachuras = []
        if not HAS_EZDXF or not os.path.exists(caminho_dxf):
            return hachuras
        
        try:
            dwg = ezdxf.readfile(caminho_dxf)
            mspace = dwg.modelspace()
            
            for entity in mspace:
                try:
                    if entity.dxftype() == 'HATCH':
                        camada = entity.dxf.layer
                        padrao = entity.dxf.pattern_name
                        x, y, bbox = self.obter_coordenadas_hachura(entity)
                        poly_real = self.extrair_poligono_hachura(entity)
                        
                        if x != 0 or y != 0:
                            hachuras.append({
                                'camada': camada,
                                'padrao': padrao,
                                'x': float(x),
                                'y': float(y),
                                'bbox': bbox,
                                'poly_real': poly_real,
                                'nome': f"HATCH_{len(hachuras)}",
                                'tipo': 'HATCH'
                            })
                except Exception:
                    continue
            return hachuras
        except Exception as e:
            st.warning(f"⚠️ Erro ao carregar hachuras: {e}")
            return []
    
    def criar_boundary_hachuras(self):
        """Converte as hachuras do DXF em objetos georreferenciados do Shapely."""
        if not HAS_SHAPELY:
            return
        
        self.hachuras_com_boundary = []
        
        for hachura in self.hachuras_filtradas:
            poly_real = hachura.get('poly_real')
            tipo_boundary = ""
            polygon = None
            
            # Prioriza a geometria EXATA (Paths)
            if poly_real and isinstance(poly_real, Polygon):
                polygon = poly_real
                tipo_boundary = "Geometria Exata"
            else:
                # Fallback: Retângulo delimitador (BBox)
                bbox = hachura.get('bbox', {})
                xmin, xmax = bbox.get('xmin', 0), bbox.get('xmax', 0)
                ymin, ymax = bbox.get('ymin', 0), bbox.get('ymax', 0)
                if xmin != 0 or xmax != 0:
                    coords = [(xmin, ymin), (xmax, ymin), (xmax, ymax), (xmin, ymax), (xmin, ymin)]
                    polygon = Polygon(coords)
                    tipo_boundary = "Retângulo (BBox)"
            
            if polygon:
                hachura_copia = hachura.copy()
                hachura_copia['boundary'] = polygon
                hachura_copia['tipo_boundary'] = tipo_boundary
                self.hachuras_com_boundary.append(hachura_copia)
        
        st.write(f"🔲 **{len(self.hachuras_com_boundary)}** boundaries de zoneamento mapeadas matematicamente.")
    
    def filtrar_hachuras_0z(self) -> List[dict]:
        """Filtra apenas as hachuras pertencentes às camadas de Zoneamento (iniciadas com 0_Z)."""
        filtradas = []
        for hachura in self.hachuras_cache:
            if hachura.get('camada', '').startswith('0_Z'):
                filtradas.append(hachura)
        return filtradas
    
    def carregar_arquivo(self, caminho_dxf: str) -> bool:
        """Pipeline principal de carregamento e indexação do DXF."""
        if self.arquivo_carregado == caminho_dxf:
            return True
        
        with st.expander("🛠️ Logs de Leitura Geométrica (DXF)", expanded=False):
            st.write(f"📥 Processando o arquivo em memória: `{os.path.basename(caminho_dxf)}`")
            
            self.textos_cache = self.carregar_textos_de_dxf(caminho_dxf)
            self.hachuras_cache = self.carregar_hachuras_de_dxf(caminho_dxf)
            self.hachuras_filtradas = self.filtrar_hachuras_0z()
            
            self.arquivo_carregado = caminho_dxf
            
            if not self.textos_cache:
                st.error("Nenhum texto utilizável foi encontrado no desenho base.")
                return False
            
            self.criar_boundary_hachuras()
            
        return True
    
    def processar_busca(self, termo1: str, termo2: str, termo3: str) -> bool:
        """Encontra todos os textos que correspondem à Rua, Lote e Quadra."""
        self.lista1, self.lista2, self.lista3 = [], [], []
        
        if not termo1 or not termo2 or not termo3:
            return False
        if not self.textos_cache:
            return False
        
        t1_busca = f"*{self.normalizar_texto(termo1)}*"
        t2_busca = f"*{self.normalizar_texto(termo2)}*"
        t3_busca = f"*{self.normalizar_texto(termo3)}*"
        
        for texto_obj in self.textos_cache:
            txt_limpo = texto_obj.get('texto_limpo', '')
            if not txt_limpo: continue
            
            pt = (texto_obj.get('x', 0), texto_obj.get('y', 0))
            ent_nome = texto_obj.get('nome', '')
            txt = texto_obj.get('texto', '')
            
            if self.wcmatch(txt_limpo, t1_busca): self.lista1.append({'ponto': pt, 'nome': ent_nome, 'texto': txt})
            if self.wcmatch(txt_limpo, t2_busca): self.lista2.append({'ponto': pt, 'nome': ent_nome, 'texto': txt})
            if self.wcmatch(txt_limpo, t3_busca): self.lista3.append({'ponto': pt, 'nome': ent_nome, 'texto': txt})
        
        if not (self.lista1 and self.lista2 and self.lista3):
            st.warning("⚠️ Os três termos solicitados (Rua, Lote e Quadra) não foram encontrados simultaneamente no DXF.")
            return False
        return True
    
    def encontrar_trio_mais_proximo(self) -> bool:
        """Avalia a distância vetorial entre as ocorrências para confirmar qual é o lote correto."""
        minDist = float('inf')
        for item2 in self.lista2:
            p2 = item2['ponto']
            # Otimização: Pega apenas as 10 ruas e quadras mais próximas para checar
            ruas_proximas = sorted(self.lista1, key=lambda x: self.distance(x['ponto'], p2))[:10]
            quadras_proximas = sorted(self.lista3, key=lambda x: self.distance(x['ponto'], p2))[:10]
            
            for item1 in ruas_proximas:
                p1 = item1['ponto']
                for item3 in quadras_proximas:
                    p3 = item3['ponto']
                    distAtual = self.distance(p1, p2) + self.distance(p1, p3) + self.distance(p2, p3)
                    
                    if distAtual < minDist:
                        minDist = distAtual
                        self.melhor1, self.melhor2, self.melhor3 = p1, p2, p3
                        self.melhorEnt = item2['nome']
        
        if not self.melhor2:
            return False
        return True
    
    def buscar_zona(self, coordX: float, coordY: float) -> bool:
        """Cruza as coordenadas (X, Y) encontradas com os limites dos polígonos de zona."""
        if not HAS_SHAPELY or not self.hachuras_com_boundary:
            self.zonaFinal = "ZH2"
            return False
        
        ponto_lote = Point(coordX, coordY)
        candidatos = []
        
        for hachura in self.hachuras_com_boundary:
            boundary = hachura.get('boundary')
            camada = hachura.get('camada')
            
            if boundary and boundary.contains(ponto_lote):
                candidatos.append({
                    'camada': camada,
                    'area': boundary.area
                })
        
        if candidatos:
            # Ordena pela menor área para evitar falsos positivos com grandes bounds engolindo bounds pequenos.
            candidatos.sort(key=lambda x: x['area'])
            vencedor = candidatos[0]
            
            # Limpeza do nome da camada para extração exata da sigla de zona
            zona_limpa = vencedor['camada'].replace("0_ZONA_", "").replace("0_Z_", "").strip()
            self.zonaFinal = zona_limpa if zona_limpa else vencedor['camada']
            return True
        else:
            self.zonaFinal = "ZH2"
            return False
    
    def iniciar(self, caminho_dxf: str) -> bool:
        """Prepara o arquivo (Chamado pela interface principal)."""
        if not os.path.exists(caminho_dxf):
            st.error(f"❌ Arquivo CAD temp não foi encontrado: {caminho_dxf}")
            return False
        return self.carregar_arquivo(caminho_dxf)

    def busca(self, termo_rua: str, termo_lote: str, termo_quadra: str) -> dict:
        """Recebe os dados extraídos pelo LLM e devolve a zona calculada."""
        st.write(f"🗺️ Triangulando coordenadas para: Rua `{termo_rua}`, Lote `{termo_lote}`, Quadra `{termo_quadra}`")
        
        if not self.processar_busca(termo_rua, termo_lote, termo_quadra): 
            return {'sucesso': False, 'zona_final': 'ZH2'}
            
        if not self.encontrar_trio_mais_proximo(): 
            return {'sucesso': False, 'zona_final': 'ZH2'}
        
        coordX, coordY = self.melhor2
        self.buscar_zona(coordX, coordY)
        
        return {'sucesso': True, 'zona_final': self.zonaFinal}