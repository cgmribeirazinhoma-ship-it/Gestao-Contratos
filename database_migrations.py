import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'gestao_contratos.db')

def migrar_banco():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # 1. Obter colunas existentes em 'orgaos'
    cursor.execute("PRAGMA table_info(orgaos)")
    colunas_orgaos = [col[1] for col in cursor.fetchall()]

    novas_colunas_orgaos = {
        'cnpj': "TEXT DEFAULT '01.612.834/0001-86'",
        'cep': "TEXT DEFAULT '65928-000'",
        'endereco': "TEXT DEFAULT 'Rua Principal, s/n'",
        'bairro': "TEXT DEFAULT 'Centro'",
        'cidade': "TEXT DEFAULT 'Ribeirãozinho do Maranhão'",
        'uf': "TEXT DEFAULT 'MA'",
        'secretario_padrao': "TEXT",
        'cargo_secretario': "TEXT DEFAULT 'Secretário(a) Municipal Titular'",
        'fiscal_padrao': "TEXT",
        'dotacao_padrao': "TEXT",
        'telefone': "TEXT",
        'email': "TEXT",
        'ativo': "INTEGER DEFAULT 1"
    }

    for col, def_tipo in novas_colunas_orgaos.items():
        if col not in colunas_orgaos:
            cursor.execute(f"ALTER TABLE orgaos ADD COLUMN {col} {def_tipo}")
            print(f"Coluna '{col}' adicionada à tabela 'orgaos'.")

    # 2. Obter colunas existentes em 'contratos'
    cursor.execute("PRAGMA table_info(contratos)")
    colunas_contratos = [col[1] for col in cursor.fetchall()]

    novas_colunas_contratos = {
        'modelo_agu': "TEXT DEFAULT 'COMPRAS'",
        'dotacao_orcamentaria': "TEXT",
        'chave_agu': "TEXT",
        'cep_orgao': "TEXT",
        'endereco_orgao': "TEXT",
        'bairro_orgao': "TEXT",
        'cidade_orgao': "TEXT DEFAULT 'Ribeirãozinho do Maranhão'",
        'uf_orgao': "TEXT DEFAULT 'MA'",
        'cnpj_orgao': "TEXT DEFAULT '01.612.834/0001-86'"
    }

    for col, def_tipo in novas_colunas_contratos.items():
        if col not in colunas_contratos:
            cursor.execute(f"ALTER TABLE contratos ADD COLUMN {col} {def_tipo}")
            print(f"Coluna '{col}' adicionada à tabela 'contratos'.")

    # Índice único para garantir que a numeração seja contínua sem repetição por ano
    cursor.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_contratos_num_ano ON contratos (numero_contrato, ano_contrato)")

    # 3. Popular/Atualizar dados padrão dos órgãos conhecidos
    mapa_orgaos_dados = {
        "SECRETARIA MUNICIPAL DE EDUCAÇÃO": {
            "secretario": "GERALDO EVANDRO BRAGA DE SOUSA",
            "cargo": "Secretário Municipal de Educação",
            "endereco": "Av. Tancredo Neves, nº 120",
            "bairro": "Centro",
            "cep": "65928-000",
            "fiscal": "CHARLIANE DE ABREU MACIEL",
            "dotacao": "Fundo Municipal de Educação / SEMED - 02.05.00 - 12.361.0402.2022.0000"
        },
        "SECRETARIA MUNICIPAL DE SAÚDE": {
            "secretario": "SIRLEIDE MARINHO DOS SANTOS",
            "cargo": "Secretária Municipal de Saúde",
            "endereco": "Rua do Comércio, s/n",
            "bairro": "Centro",
            "cep": "65928-000",
            "fiscal": "JOSIVAN SILVA AROUCHA"
        },
        "SECRETARIA MUNICIPAL DE FINANÇAS, FAZENDA E RECEITA": {
            "secretario": "DANIEL SILVA PEREIRA",
            "cargo": "Secretário Municipal de Finanças",
            "endereco": "Rua Principal, nº 45",
            "bairro": "Centro",
            "cep": "65928-000",
            "fiscal": "WILSON FERREIRA SOARES"
        },
        "SECRETARIA MUNICIPAL DE ASSISTÊNCIA SOCIAL": {
            "secretario": "FERNANDA NUNES ROCHA",
            "cargo": "Secretária Municipal de Assistência Social",
            "endereco": "Rua São Luís, nº 80",
            "bairro": "Centro",
            "cep": "65928-000",
            "fiscal": "Fiscal de Assistência Social"
        },
        "SECRETARIA MUNICIPAL DE CULTURA E TURISMO": {
            "secretario": "ELANDIAS BEZERRA SOUSA",
            "cargo": "Secretário Municipal de Cultura e Turismo",
            "endereco": "Praça Central, s/n",
            "bairro": "Centro",
            "cep": "65928-000",
            "fiscal": "Fiscal de Cultura"
        },
        "SECRETARIA MUNICIPAL DE ADMINISTRAÇÃO": {
            "secretario": "GUSTAVO ANDRADE",
            "cargo": "Secretário Municipal de Administração",
            "endereco": "Rua Principal, s/n",
            "bairro": "Centro",
            "cep": "65928-000",
            "fiscal": "Fiscal de Administração"
        },
        "SECRETARIA MUNICIPAL DE INFRAESTRUTURA E OBRAS": {
            "secretario": "GUSTAVO ANDRADE",
            "cargo": "Secretário Municipal de Infraestrutura e Obras",
            "endereco": "Av. Brasil, nº 200",
            "bairro": "Industrial",
            "cep": "65928-000",
            "fiscal": "Engenheiro Fiscal de Obras"
        },
        "SERVIÇO AUTÔNOMO DE ÁGUA E ESGOTO - SAAE": {
            "secretario": "DIRETOR GERAL SAAE",
            "cargo": "Diretor Geral do SAAE",
            "endereco": "Rua da Paz, nº 50",
            "bairro": "Centro",
            "cep": "65928-000",
            "fiscal": "Fiscal Operacional SAAE"
        },
        "ASSESSORIA JURÍDICA E GABINETE": {
            "secretario": "PROCURADORIA GERAL DO MUNICÍPIO",
            "cargo": "Procurador Geral do Município",
            "endereco": "Palácio Municipal, Rua Principal, s/n",
            "bairro": "Centro",
            "cep": "65928-000",
            "fiscal": "Assessor de Gabinete"
        }
    }

    for nome_orgao, dados in mapa_orgaos_dados.items():
        cursor.execute("""
            UPDATE orgaos 
            SET secretario_padrao = COALESCE(secretario_padrao, :sec),
                cargo_secretario = COALESCE(cargo_secretario, :carg),
                endereco = COALESCE(endereco, :end),
                bairro = COALESCE(bairro, :bai),
                cep = COALESCE(cep, :cep),
                fiscal_padrao = COALESCE(fiscal_padrao, :fisc),
                cidade = 'Ribeirãozinho do Maranhão',
                uf = 'MA'
            WHERE nome = :nome OR nome LIKE :nome_like
        """, {
            "sec": dados["secretario"],
            "carg": dados["cargo"],
            "end": dados["endereco"],
            "bai": dados["bairro"],
            "cep": dados["cep"],
            "fisc": dados["fiscal"],
            "nome": nome_orgao,
            "nome_like": f"%{nome_orgao[:15]}%"
        })

    # 4. Criar tabela de Termos Aditivos caso não exista
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS termos_aditivos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            contrato_id INTEGER NOT NULL,
            numero_aditivo INTEGER NOT NULL,
            numero_completo TEXT NOT NULL,
            ano_aditivo INTEGER NOT NULL,
            tipo_aditivo TEXT NOT NULL,
            regime_legal TEXT DEFAULT 'LEI_14133_2021',
            processo_adm TEXT,
            data_assinatura DATE,
            data_publicacao DATE,
            nova_data_vencimento DATE,
            prazo_aditado_meses INTEGER DEFAULT 0,
            valor_aditado REAL DEFAULT 0.0,
            percentual_aditado REAL DEFAULT 0.0,
            novo_valor_total REAL,
            objeto_aditivo TEXT NOT NULL,
            justificativa TEXT,
            dotacao_orcamentaria TEXT,
            garantia_execucao TEXT,
            secretario TEXT,
            fiscal TEXT,
            status TEXT DEFAULT 'ATIVO',
            criado_por INTEGER,
            data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (contrato_id) REFERENCES contratos(id) ON DELETE CASCADE,
            FOREIGN KEY (criado_por) REFERENCES usuarios(id)
        )
    """)
    print("Tabela 'termos_aditivos' verificada/criada.")

    # 5. Obter colunas existentes em 'fornecedores' e adicionar campos completos da Receita Federal
    cursor.execute("PRAGMA table_info(fornecedores)")
    colunas_fornec = [col[1] for col in cursor.fetchall()]

    novas_colunas_fornec = {
        'numero': "TEXT DEFAULT 's/n'",
        'complemento': "TEXT",
        'bairro': "TEXT DEFAULT 'Centro'",
        'cidade': "TEXT DEFAULT 'Imperatriz'",
        'uf': "TEXT DEFAULT 'MA'",
        'cep': "TEXT DEFAULT '65900-000'",
        'cargo_representante': "TEXT DEFAULT 'Sócio Administrador'",
        'situacao_cadastral': "TEXT DEFAULT 'ATIVA'"
    }

    for col, def_tipo in novas_colunas_fornec.items():
        if col not in colunas_fornec:
            cursor.execute(f"ALTER TABLE fornecedores ADD COLUMN {col} {def_tipo}")
            print(f"Coluna '{col}' adicionada à tabela 'fornecedores'.")

    # 6. Higienização e Padronização Automática de CNPJ/CPF (padrão com pontos e barra)
    import re
    def fmt_cnpj_cpf(val):
        if not val:
            return val
        s = re.sub(r'\D', '', str(val)).strip()
        if len(s) == 14:
            return f"{s[:2]}.{s[2:5]}.{s[5:8]}/{s[8:12]}-{s[12:]}"
        elif len(s) == 11:
            return f"{s[:3]}.{s[3:6]}.{s[6:9]}-{s[9:]}"
        return str(val).strip()

    cursor.execute("SELECT id, cnpj_cpf FROM fornecedores WHERE cnpj_cpf IS NOT NULL")
    for fid, cnpj in cursor.fetchall():
        fmt = fmt_cnpj_cpf(cnpj)
        if fmt != cnpj:
            cursor.execute("UPDATE fornecedores SET cnpj_cpf = ? WHERE id = ?", (fmt, fid))

    cursor.execute("SELECT id, cnpj_orgao FROM contratos WHERE cnpj_orgao IS NOT NULL")
    for cid, cnpj in cursor.fetchall():
        fmt = fmt_cnpj_cpf(cnpj)
        if fmt != cnpj:
            cursor.execute("UPDATE contratos SET cnpj_orgao = ? WHERE id = ?", (fmt, cid))

    cursor.execute("SELECT id, cnpj FROM orgaos WHERE cnpj IS NOT NULL")
    for oid, cnpj in cursor.fetchall():
        fmt = fmt_cnpj_cpf(cnpj)
        if fmt != cnpj:
            cursor.execute("UPDATE orgaos SET cnpj = ? WHERE id = ?", (fmt, oid))

    conn.commit()
    conn.close()

    print("Migração e padronização de dados concluídas com sucesso!")

if __name__ == '__main__':
    migrar_banco()
