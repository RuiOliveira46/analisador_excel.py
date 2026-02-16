"""
Gerador de ficheiro Excel de exemplo
Cria dados de vendas fictícios para testar o analisador
"""

import pandas as pd
from datetime import datetime, timedelta
import random

# Configurar dados de exemplo
cidades = ['Lisboa', 'Porto', 'Coimbra', 'Braga', 'Faro']
produtos = ['Laptop', 'Telemóvel', 'Tablet', 'Monitor', 'Teclado', 'Rato']
vendedores = ['Ana Silva', 'João Santos', 'Maria Costa', 'Pedro Alves', 'Sofia Oliveira']

# Gerar dados
dados = []
data_inicial = datetime(2024, 1, 1)

for i in range(200):
    dados.append({
        'ID': i + 1,
        'Data': data_inicial + timedelta(days=random.randint(0, 365)),
        'Produto': random.choice(produtos),
        'Quantidade': random.randint(1, 10),
        'Preco_Unitario': round(random.uniform(50, 2000), 2),
        'Cidade': random.choice(cidades),
        'Vendedor': random.choice(vendedores),
        'Cliente': f"Cliente_{random.randint(1, 50)}"
    })

# Calcular valor total
df = pd.DataFrame(dados)
df['Valor_Total'] = df['Quantidade'] * df['Preco_Unitario']
df['Valor_Total'] = df['Valor_Total'].round(2)

# Guardar ficheiro
nome_ficheiro = 'vendas_exemplo.xlsx'
df.to_excel(nome_ficheiro, index=False)

print(f"✅ Ficheiro '{nome_ficheiro}' criado com sucesso!")
print(f"📊 {len(df)} registos de vendas gerados")
print("\nPrimeiras linhas:")
print(df.head(10))
