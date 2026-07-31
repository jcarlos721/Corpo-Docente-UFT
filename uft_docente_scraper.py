import requests
import pandas as pd
import uuid
import time
from datetime import datetime
from urllib.parse import urlparse

# ==========================================
# CONFIGURAÇÕES DA API
# ==========================================
BASE_URL = "https://www.uft.edu.br"
# O endpoint @search é poderoso. Ele busca todas as páginas com o ID "corpo-docente" no site todo.
SEARCH_URL = f"{BASE_URL}/++api++/@search?id=corpo-docente&b_size=1000"
HEADERS = {"Accept": "application/json"}

def extrair_hierarquia_da_url(url):
    """
    Função para modelar o Banco de Dados.
    Exemplo de URL: https://www.uft.edu.br/campus/arraias/cursos/graduacao/direito/corpo-docente
    Isolamos: campus (arraias), modalidade (graduacao), curso (direito).
    """
    partes = urlparse(url).path.strip('/').split('/')
    
    # Valores padrão caso a URL tenha uma estrutura diferente
    campus = "Desconhecido"
    modalidade = "Desconhecida"
    curso = "Desconhecido"
    
    try:
        if 'campus' in partes:
            idx_campus = partes.index('campus')
            campus = partes[idx_campus + 1].capitalize()
        
        if 'cursos' in partes:
            idx_cursos = partes.index('cursos')
            modalidade = partes[idx_cursos + 1].capitalize()
            curso = partes[idx_cursos + 2].replace('-', ' ').title()
    except IndexError:
        pass # Mantém os valores desconhecidos se a URL for fora do padrão
        
    return campus, modalidade, curso

def coletar_professores():
    print("🔍 Buscando todas as páginas de Corpo Docente no sistema da UFT...")
    resposta_busca = requests.get(SEARCH_URL, headers=HEADERS)
    
    if resposta_busca.status_code != 200:
        print("❌ Erro ao acessar a busca da UFT.")
        return

    itens_encontrados = resposta_busca.json().get('items', [])
    print(f"✅ Foram encontradas {len(itens_encontrados)} páginas de Corpo Docente.")
    
    dados_modelados = []
    data_extracao_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # ==========================================
    # VARREDURA (CRAWLER)
    # ==========================================
    for index, item in enumerate(itens_encontrados):
        url_original = item.get('@id')
        
        # Filtro de segurança: garantir que estamos pegando dados de cursos
        if 'cursos' not in url_original:
            continue
            
        campus, modalidade, curso = extrair_hierarquia_da_url(url_original)
        
        # Converte a URL da página normal para a URL da API oculta injetando ++api++
        api_url = url_original.replace("https://www.uft.edu.br/", "https://www.uft.edu.br/++api++/")
        
        print(f"[{index+1}/{len(itens_encontrados)}] Extraindo: {campus} > {curso}...")
        
        try:
            resp_pagina = requests.get(api_url, headers=HEADERS)
            if resp_pagina.status_code != 200:
                continue
                
            json_pagina = resp_pagina.json()
            blocos = json_pagina.get("blocks", {})
            
            for id_bloco, conteudo in blocos.items():
                if conteudo.get("@type") == "profileBlock":
                    pessoas = conteudo.get("pessoas", [])
                    
                    for prof in pessoas:
                        # ----------------------------------------------------
                        # MODELAGEM DE DADOS PARA O FUTURO BANCO (RDBMS)
                        # ----------------------------------------------------
                        registro = {
                            "id_registro_bd": str(uuid.uuid4()),  # Chave Primária Universal (PK)
                            "id_professor_plone": prof.get("@id", ""), # Referência externa
                            "nome": prof.get("nome", "Não informado").strip(),
                            "cargo": prof.get("cargo", "Sem cargo").strip(),
                            "email": prof.get("email", "").strip(),
                            "lattes_url": prof.get("lattes", "").strip(),
                            "campus_nome": campus, # Futura Chave Estrangeira (FK) para tabela 'Campus'
                            "curso_nome": curso,   # Futura Chave Estrangeira (FK) para tabela 'Cursos'
                            "modalidade": modalidade,
                            "url_fonte_dados": url_original,
                            "data_extracao": data_extracao_atual
                        }
                        dados_modelados.append(registro)
                        
        except Exception as e:
            print(f"Erro ao processar {curso} em {campus}: {e}")
            
        # Pausa de meio segundo para não sobrecarregar o servidor da universidade (Boas práticas)
        time.sleep(0.5)

    # ==========================================
    # EXPORTAÇÃO E GERAÇÃO DA PLANILHA EXCEL
    # ==========================================
    if dados_modelados:
        print(f"\n🚀 Varredura concluída! {len(dados_modelados)} professores encontrados. Gerando Excel...")
        
        # O Pandas transforma nossa lista de dicionários em uma tabela instantaneamente
        df = pd.DataFrame(dados_modelados)
        
        # Ordenamos a planilha automaticamente por Campus e depois por Curso
        df = df.sort_values(by=['campus_nome', 'curso_nome', 'nome'])
        
        nome_arquivo = "Base_Dados_Docentes_UFT.xlsx"
        
        # Salva em formato Excel moderno (.xlsx)
        df.to_excel(nome_arquivo, index=False, engine='openpyxl')
        
        print(f"🎉 Sucesso! Arquivo salvo como '{nome_arquivo}'.")
        print("💡 Dica: Esta planilha já está pronta para ser importada em qualquer Banco de Dados (SQL).")
    else:
        print("Nenhum dado válido foi encontrado.")

if __name__ == "__main__":
    coletar_professores()