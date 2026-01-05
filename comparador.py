import os
import re
import subprocess
import pandas as pd
from docx import Document
from lxml import etree
from odf.opendocument import load
from odf import teletype, text

# ==============================================================================
# 1. UTILITÁRIOS
# ==============================================================================

def normalizar_chave(texto):
    """Remove caracteres especiais para facilitar o cruzamento de dados."""
    t = texto.lower()
    return re.sub(r'[^\w]', '', t)

def analisar_campo(texto, is_colorido):
    """
    Define Tipo e Classificação.
    Lógica: Se tem cor definida (diferente de preto) OU Estilo de destaque -> Opcional.
    """
    tipos = []
    
    # Regra 1: Preenchimento ([], XX, <>, (...))
    if re.search(r"\[.*?\]|XX|<.*?>|\(\.\.\.\)", texto):
        tipos.append("Preenchimento")
    
    # Regra 2: Alternativa (OU isolado)
    if re.search(r"\bOU\b", texto):
        tipos.append("Alternativa")
        
    # Regra 3: Classificação por Cor/Estilo
    classificacao = "Obrigatório"
    if is_colorido:
        classificacao = "Opcional"
        tipos.append("Escolha")
        
    tipo_final = ", ".join(tipos) if tipos else "Texto Fixo"
    return tipo_final, classificacao

# ==============================================================================
# 2. MOTOR DE COR & ESTILO (XML PROFUNDO)
# ==============================================================================

def verificar_indicativo_de_cor_ou_estilo(paragrafo_element):
    """
    Varre o XML buscando:
    1. Cores explícitas diferentes de preto/auto.
    2. Realces (Highlight) ou Sombreamento (Shading).
    3. Nomes de ESTILOS que sugiram cor (ex: 'Texto Vermelho', 'Emphasis').
    """
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    palavras_chave_estilo = ['VERMELHO', 'RED', 'COLORIDO', 'DESTAQUE', 'EMPHASIS', 'OPCIONAL', 'ALERT', 'OBSERVAÇÃO']

    try:
        # A. Checa Estilo do Parágrafo (<w:pPr><w:pStyle ...>)
        pPr = paragrafo_element.find('w:pPr', ns)
        if pPr is not None:
            pStyle = pPr.find('w:pStyle', ns)
            if pStyle is not None:
                style_val = pStyle.get(f"{{{ns['w']}}}val", "").upper()
                if any(k in style_val for k in palavras_chave_estilo):
                    return True

        # B. Checa Runs (trechos de texto)
        runs = paragrafo_element.findall('.//w:r', ns)
        
        for run in runs:
            rPr = run.find('w:rPr', ns)
            if rPr is not None:
                # 1. Checa Estilo do Run (<w:rStyle ...>)
                rStyle = rPr.find('w:rStyle', ns)
                if rStyle is not None:
                    style_val = rStyle.get(f"{{{ns['w']}}}val", "").upper()
                    if any(k in style_val for k in palavras_chave_estilo):
                        return True

                # 2. Checa Cor Explícita (<w:color>)
                color_tag = rPr.find('w:color', ns)
                if color_tag is not None:
                    val = color_tag.get(f"{{{ns['w']}}}val")
                    theme = color_tag.get(f"{{{ns['w']}}}themeColor")
                    
                    if theme: return True # Tem cor de tema
                    if val and val.lower() not in ['000000', 'auto']: return True # Tem cor Hex

                # 3. Checa Realce/Highlight
                highlight = rPr.find('w:highlight', ns)
                if highlight is not None:
                    val = highlight.get(f"{{{ns['w']}}}val")
                    if val and val.lower() != 'none': return True

                # 4. Checa Sombreamento/Fundo
                shd = rPr.find('w:shd', ns)
                if shd is not None:
                    fill = shd.get(f"{{{ns['w']}}}fill")
                    if fill and fill.lower() not in ['auto', 'ffffff', '000000']: return True

    except:
        pass
    
    return False

# ==============================================================================
# 3. MOTOR HÍBRIDO (DOCX)
# ==============================================================================

def converter_docx_para_txt(caminho_docx):
    """LibreOffice: Extrai índices reais (1.1)."""
    try:
        subprocess.run(
            ['soffice', '--headless', '--convert-to', 'txt:Text', '--outdir', '.', caminho_docx],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE
        )
        nome_txt = os.path.splitext(os.path.basename(caminho_docx))[0] + ".txt"
        return nome_txt if os.path.exists(nome_txt) else None
    except:
        return None

def mapear_indices_txt(caminho_txt):
    mapa = {}
    try:
        with open(caminho_txt, 'r', encoding='utf-8') as f:
            linhas = f.readlines()
    except: return {}

    for linha in linhas:
        linha = linha.strip()
        if not linha: continue
        
        match = re.match(r"^([\d\.]+)([\s\.\-\)]+)(.*)", linha)
        if match:
            indice = match.group(1).strip()
            texto = match.group(3).strip()
            if texto:
                k = normalizar_chave(texto)
                if k not in mapa: mapa[k] = []
                mapa[k].append(indice)
    return mapa

def extrair_comentarios_docx(doc):
    comentarios = {}
    try:
        for part in doc.part.package.parts:
            if part.partname.endswith('comments.xml'):
                root = etree.fromstring(part.blob)
                ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                for c in root.findall('.//w:comment', ns):
                    cid = c.get(f"{{{ns['w']}}}id")
                    ctext = "".join([n.text for n in c.findall('.//w:t', ns) if n.text])
                    if cid and ctext: comentarios[cid] = ctext
                break
    except: pass
    return comentarios

def obter_nota_docx(paragrafo, mapa_comentarios):
    notas = []
    try:
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        for ref in paragrafo._element.findall('.//w:commentReference', ns):
            cid = ref.get(f"{{{ns['w']}}}id")
            if cid in mapa_comentarios: notas.append(mapa_comentarios[cid])
    except: pass
    return " | ".join(notas)

def processar_docx_hibrido(caminho_arquivo):
    print(f"Lendo DOCX: {caminho_arquivo}")
    
    # 1. Índices (TXT)
    caminho_txt = converter_docx_para_txt(caminho_arquivo)
    mapa_indices = mapear_indices_txt(caminho_txt) if caminho_txt else {}
    
    # 2. Metadados (DOCX)
    try:
        doc = Document(caminho_arquivo)
    except:
        print("[ERRO] Arquivo inválido.")
        return []
    
    mapa_comentarios = extrair_comentarios_docx(doc)
    dados = []
    
    print("   > Mapeando campos...")

    for p in doc.paragraphs:
        texto = p.text.strip()
        if not texto: continue
        
        # Índice
        chave = normalizar_chave(texto)
        indice = ""
        if chave in mapa_indices and mapa_indices[chave]:
            indice = mapa_indices[chave].pop(0)
        
        # Metadados (Verifica Cor + Estilo)
        eh_colorido = verificar_indicativo_de_cor_ou_estilo(p._element)
        
        nota = obter_nota_docx(p, mapa_comentarios)
        tipo, classif = analisar_campo(texto, eh_colorido)
        
        # Filtro de ruído
        if not indice and len(texto) < 3: continue

        dados.append({
            "indice": indice,
            "texto": texto,
            "tipo": tipo,
            "classificacao": classif,
            "nota": nota
        })

    if caminho_txt and os.path.exists(caminho_txt):
        os.remove(caminho_txt)
        
    return dados

# ==============================================================================
# 3. MOTOR ODT
# ==============================================================================

def processar_odt_padrao(caminho_arquivo):
    print(f"Lendo ODT: {caminho_arquivo}")
    try:
        textdoc = load(caminho_arquivo)
        paragrafos = textdoc.getElementsByType(text.P)
    except: return []
    dados = []
    for p in paragrafos:
        linha = teletype.extractText(p).strip()
        if not linha: continue
        match = re.match(r"^([\d\.]+)([\s\.\-\)]+)(.*)", linha)
        if match:
            idx = match.group(1).strip()
            txt = match.group(3).strip()
        else:
            idx = ""
            txt = linha
        tipo, classif = analisar_campo(txt, False)
        if not idx and len(txt) < 3: continue
        dados.append({"indice": idx, "texto": txt, "tipo": tipo, "classificacao": classif, "nota": ""})
    return dados

# ==============================================================================
# 4. ORQUESTRADOR
# ==============================================================================

def selecionar_arquivo(msg):
    arqs = [f for f in os.listdir('.') if f.lower().endswith(('.docx', '.odt')) and not f.startswith('~')]
    arqs.sort()
    if not arqs:
        print("\n[ERRO] Nenhum arquivo compatível.")
        return None
    print(f"\n--- {msg} ---")
    for i, a in enumerate(arqs): print(f"[{i+1}] {a}")
    while True:
        try:
            x = int(input("Número: ")) - 1
            if 0 <= x < len(arqs): return arqs[x]
        except: pass

def main():
    print("=== MAPEADOR HÍBRIDO V4 (Detecção de Estilos) ===")
    
    arquivo_base = selecionar_arquivo("DOCUMENTO PRINCIPAL")
    if not arquivo_base: return
    
    comparar = False
    arquivo_comp = None
    resp = input("\nComparar com outro arquivo? (S/N): ").upper().strip()
    if resp == 'S':
        comparar = True
        arquivo_comp = selecionar_arquivo("SEGUNDO DOCUMENTO")
        if not arquivo_comp: return

    # Base
    if arquivo_base.lower().endswith('.docx'):
        dados_base = processar_docx_hibrido(arquivo_base)
    else:
        dados_base = processar_odt_padrao(arquivo_base)

    if not dados_base: return

    # Comparação
    mapa_comp = {}
    if comparar and arquivo_comp:
        print(f"Lendo Comparação: {arquivo_comp}")
        if arquivo_comp.lower().endswith('.docx'):
            raw_comp = processar_docx_hibrido(arquivo_comp)
        else:
            raw_comp = processar_odt_padrao(arquivo_comp)
        
        for item in raw_comp:
            k = normalizar_chave(item['texto'])
            mapa_comp[k] = (item['indice'], item['texto'])

    # Geração
    lista_final = []
    # Usando o nome completo do arquivo
    nome_base = os.path.splitext(arquivo_base)[0]
    nome_comp = os.path.splitext(arquivo_comp)[0] if arquivo_comp else ""

    for item in dados_base:
        linha = {
            f"ID ({nome_base})": item['indice'],
            f"Rótulo ({nome_base})": item['texto'],
            "Tipo de Campo": item['tipo'],
            "Classificação": item['classificacao'],
            "Nota Explicativa": item['nota']
        }
        
        if comparar:
            k = normalizar_chave(item['texto'])
            if k in mapa_comp:
                idx_c, txt_c = mapa_comp[k]
                linha[f"ID ({nome_comp})"] = idx_c
                linha[f"Presente em {nome_comp}?"] = "Sim"
            else:
                linha[f"ID ({nome_comp})"] = "X"
                linha[f"Presente em {nome_comp}?"] = "Não"
        
        lista_final.append(linha)

    if lista_final:
        df = pd.DataFrame(lista_final)
        if comparar:
            nome_saida = f"Mapeamento_{nome_base}_VS_{nome_comp}.xlsx"
        else:
            nome_saida = f"Mapeamento_{nome_base}.xlsx"
        try:
            df.to_excel(nome_saida, index=False)
            print(f"\n[SUCESSO] Salvo como: {nome_saida}")
            print(f"Total: {len(df)}")
        except:
            print("[ERRO] Falha ao salvar arquivo (verifique se está aberto).")
    else:
        print("[AVISO] Nenhum dado encontrado.")

if __name__ == "__main__":
    main()
