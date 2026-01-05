import pandas as pd
import random
from datetime import datetime, timedelta

def gerar_base_dados():
    """
    Gera uma base de dados fictícia para simular um sistema ERP.
    - 320 Clientes
    - 100 Produtos
    - 1350 Registros de Vendas
    - 270 Blocos (5 registros por bloco)
    """
    
    print("Iniciando geração de dados...")

    # --- 1. GERAR 320 CLIENTES ---
    nomes_base = ["João", "Maria", "José", "Ana", "Carlos", "Paula", "Lucas", "Beatriz", "Marcos", "Fernanda", "Ricardo", "Juliana", "Gabriel", "Larissa", "Roberto", "Camila"]
    sobrenomes_base = ["Silva", "Santos", "Oliveira", "Souza", "Pereira", "Costa", "Ferreira", "Rodrigues", "Almeida", "Nascimento", "Melo", "Barbosa", "Cardoso", "Teixeira", "Ribeiro", "Vieira"]
    
    lista_clientes = []
    for i in range(1, 321):
        nome = random.choice(nomes_base)
        sn1 = random.choice(sobrenomes_base)
        sn2 = random.choice(sobrenomes_base)
        nome_completo = f"{nome} {sn1} {sn2}"
        
        lista_clientes.append({
            "ID Cliente": i,
            "Nome": nome_completo,
            "CPF": f"{random.randint(100,999)}.{random.randint(100,999)}.{random.randint(100,999)}-{random.randint(10,99)}",
            "Email": f"{nome.lower()}.{sn1.lower()}{i}@exemplo.com",
            "Telefone": f"(11) 9{random.randint(7000, 9999)}-{random.randint(1000, 9999)}"
        })
    
    df_clientes = pd.DataFrame(lista_clientes)

    # --- 2. GERAR 100 PRODUTOS ---
    categorias = ["Eletrônicos", "Móveis", "Acessórios", "Informática", "Eletrodomésticos"]
    objetos = ["Cadeira", "Mesa", "Monitor", "Teclado", "Mouse", "Smartphone", "Cabo", "Suporte", "Luminária", "Headset"]
    
    lista_produtos = []
    for i in range(1, 101):
        prod_nome = f"{random.choice(objetos)} {random.choice(['Pro', 'Plus', 'Premium', 'Basic', 'Ultra'])} {i}"
        lista_produtos.append({
            "ID Produto": i,
            "Nome": prod_nome,
            "Categoria": random.choice(categorias),
            "Preco": round(random.uniform(50.0, 5000.0), 2),
            "Estoque": random.randint(10, 500)
        })
    
    df_produtos = pd.DataFrame(lista_produtos)

    # --- 3. GERAR 1350 VENDAS EM 270 BLOCOS (5 vendas por bloco) ---
    lista_vendas = []
    data_base = datetime(2023, 1, 1)
    
    id_venda_global = 1
    # 270 blocos de 5 registros = 1350 registros total
    for num_bloco in range(1, 271):
        bloco_codigo = f"BL-{num_bloco:03d}"
        
        for _ in range(5):
            prod = random.choice(lista_produtos)
            qtd = random.randint(1, 10)
            cliente = random.choice(lista_clientes)
            # Data aleatória dentro do ano de 2023
            data_venda = (data_base + timedelta(days=random.randint(0, 364))).strftime("%d/%m/%Y")
            
            lista_vendas.append({
                "ID Venda": id_venda_global,
                "Bloco": bloco_codigo,
                "Data": data_venda,
                "ID Cliente": cliente["ID Cliente"],
                "Nome Cliente": cliente["Nome"],
                "ID Produto": prod["ID Produto"],
                "Produto": prod["Nome"],
                "Quantidade": qtd,
                "Preço Unitário": prod["Preco"],
                "Total": round(prod["Preco"] * qtd, 2)
            })
            id_venda_global += 1

    df_vendas = pd.DataFrame(lista_vendas)

    # --- 4. EXPORTAR PARA EXCEL ---
    nome_arquivo = "base_erp_completa.xlsx"
    with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
        df_clientes.to_excel(writer, sheet_name='Clientes', index=False)
        df_produtos.to_excel(writer, sheet_name='Produtos', index=False)
        df_vendas.to_excel(writer, sheet_name='Vendas', index=False)

    print(f"--- GERAÇÃO CONCLUÍDA ---")
    print(f"Arquivo gerado: {nome_arquivo}")
    print(f"Total de Clientes: {len(df_clientes)}")
    print(f"Total de Produtos: {len(df_produtos)}")
    print(f"Total de Registros de Venda: {len(df_vendas)}")
    print(f"Total de Blocos: {df_vendas['Bloco'].nunique()}")
    print(f"Itens por Bloco: {int(len(df_vendas) / df_vendas['Bloco'].nunique())}")

if __name__ == "__main__":
    gerar_base_dados()