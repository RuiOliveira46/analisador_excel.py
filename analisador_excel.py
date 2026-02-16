"""
Analisador e Filtrador de Ficheiros Excel
Programa interativo para análise de dados em Excel
"""

import pandas as pd
import os
from datetime import datetime


class AnalisadorExcel:
    def __init__(self):
        self.df = None
        self.nome_ficheiro = None
        
    def carregar_ficheiro(self, caminho):
        """Carrega um ficheiro Excel"""
        try:
            # Verificar se o ficheiro existe
            if not os.path.exists(caminho):
                print(f"❌ Erro: Ficheiro '{caminho}' não encontrado!")
                return False
            
            # Carregar o ficheiro
            self.df = pd.read_excel(caminho)
            self.nome_ficheiro = caminho
            print(f"✅ Ficheiro carregado com sucesso!")
            print(f"📊 {len(self.df)} linhas e {len(self.df.columns)} colunas")
            return True
        except Exception as e:
            print(f"❌ Erro ao carregar ficheiro: {e}")
            return False
    
    def mostrar_info_basica(self):
        """Mostra informação básica sobre os dados"""
        if self.df is None:
            print("⚠️  Nenhum ficheiro carregado!")
            return
        
        print("\n" + "="*60)
        print("📋 INFORMAÇÃO BÁSICA")
        print("="*60)
        print(f"\n🔢 Número de linhas: {len(self.df)}")
        print(f"🔢 Número de colunas: {len(self.df.columns)}")
        print(f"\n📝 Colunas disponíveis:")
        for i, col in enumerate(self.df.columns, 1):
            tipo = self.df[col].dtype
            print(f"  {i}. {col} ({tipo})")
        
        print(f"\n📊 Primeiras 5 linhas:")
        print(self.df.head())
        
    def mostrar_estatisticas(self):
        """Mostra estatísticas descritivas"""
        if self.df is None:
            print("⚠️  Nenhum ficheiro carregado!")
            return
        
        print("\n" + "="*60)
        print("📈 ESTATÍSTICAS DESCRITIVAS")
        print("="*60)
        print(self.df.describe())
        
        # Valores em falta
        print("\n🔍 Valores em falta por coluna:")
        missing = self.df.isnull().sum()
        for col, count in missing.items():
            if count > 0:
                print(f"  ⚠️  {col}: {count} valores em falta")
        if missing.sum() == 0:
            print("  ✅ Sem valores em falta!")
    
    def filtrar_por_coluna(self):
        """Filtra dados por valores de uma coluna"""
        if self.df is None:
            print("⚠️  Nenhum ficheiro carregado!")
            return
        
        print("\n📋 Colunas disponíveis:")
        for i, col in enumerate(self.df.columns, 1):
            print(f"  {i}. {col}")
        
        try:
            escolha = int(input("\nEscolha o número da coluna: ")) - 1
            if escolha < 0 or escolha >= len(self.df.columns):
                print("❌ Escolha inválida!")
                return
            
            coluna = self.df.columns[escolha]
            
            # Mostrar valores únicos se forem poucos
            valores_unicos = self.df[coluna].nunique()
            if valores_unicos <= 20:
                print(f"\n📊 Valores únicos em '{coluna}':")
                for val in self.df[coluna].unique():
                    count = len(self.df[self.df[coluna] == val])
                    print(f"  - {val}: {count} registos")
            
            # Tipo de filtro
            print("\nTipo de filtro:")
            print("1. Igual a")
            print("2. Maior que")
            print("3. Menor que")
            print("4. Contém texto")
            
            tipo = input("Escolha (1-4): ")
            valor = input("Valor para filtrar: ")
            
            # Aplicar filtro
            if tipo == "1":
                # Tentar converter para número se possível
                try:
                    valor_num = float(valor)
                    df_filtrado = self.df[self.df[coluna] == valor_num]
                except:
                    df_filtrado = self.df[self.df[coluna] == valor]
            elif tipo == "2":
                df_filtrado = self.df[self.df[coluna] > float(valor)]
            elif tipo == "3":
                df_filtrado = self.df[self.df[coluna] < float(valor)]
            elif tipo == "4":
                df_filtrado = self.df[self.df[coluna].astype(str).str.contains(valor, na=False)]
            else:
                print("❌ Opção inválida!")
                return
            
            print(f"\n✅ Filtro aplicado! {len(df_filtrado)} linhas encontradas.")
            print(df_filtrado)
            
            # Opção de guardar
            guardar = input("\n💾 Guardar resultados? (s/n): ").lower()
            if guardar == 's':
                nome = input("Nome do ficheiro (sem extensão): ")
                df_filtrado.to_excel(f"{nome}.xlsx", index=False)
                print(f"✅ Guardado como '{nome}.xlsx'")
                
        except Exception as e:
            print(f"❌ Erro: {e}")
    
    def agrupar_dados(self):
        """Agrupa e resume dados"""
        if self.df is None:
            print("⚠️  Nenhum ficheiro carregado!")
            return
        
        print("\n📋 Colunas disponíveis:")
        for i, col in enumerate(self.df.columns, 1):
            print(f"  {i}. {col}")
        
        try:
            # Escolher coluna para agrupar
            grupo_idx = int(input("\nAgrupar por qual coluna? (número): ")) - 1
            coluna_grupo = self.df.columns[grupo_idx]
            
            # Escolher coluna para agregar
            print("\nColunas numéricas:")
            colunas_numericas = self.df.select_dtypes(include=['number']).columns
            for i, col in enumerate(colunas_numericas, 1):
                print(f"  {i}. {col}")
            
            agregar_idx = int(input("\nAgregar qual coluna? (número): ")) - 1
            coluna_agregar = colunas_numericas[agregar_idx]
            
            # Tipo de agregação
            print("\nTipo de agregação:")
            print("1. Soma")
            print("2. Média")
            print("3. Contagem")
            print("4. Máximo")
            print("5. Mínimo")
            
            tipo = input("Escolha (1-5): ")
            
            funcoes = {
                '1': 'sum',
                '2': 'mean',
                '3': 'count',
                '4': 'max',
                '5': 'min'
            }
            
            if tipo not in funcoes:
                print("❌ Opção inválida!")
                return
            
            # Agrupar
            resultado = self.df.groupby(coluna_grupo)[coluna_agregar].agg(funcoes[tipo])
            resultado = resultado.sort_values(ascending=False)
            
            print(f"\n📊 Resultado ({funcoes[tipo]} de '{coluna_agregar}' por '{coluna_grupo}'):")
            print(resultado)
            
            # Opção de guardar
            guardar = input("\n💾 Guardar resultados? (s/n): ").lower()
            if guardar == 's':
                nome = input("Nome do ficheiro (sem extensão): ")
                resultado.to_excel(f"{nome}.xlsx")
                print(f"✅ Guardado como '{nome}.xlsx'")
                
        except Exception as e:
            print(f"❌ Erro: {e}")
    
    def exportar_colunas_especificas(self):
        """Exporta apenas colunas selecionadas"""
        if self.df is None:
            print("⚠️  Nenhum ficheiro carregado!")
            return
        
        print("\n📋 Colunas disponíveis:")
        for i, col in enumerate(self.df.columns, 1):
            print(f"  {i}. {col}")
        
        try:
            escolhas = input("\nNúmeros das colunas a exportar (separados por vírgula): ")
            indices = [int(x.strip()) - 1 for x in escolhas.split(',')]
            
            colunas_selecionadas = [self.df.columns[i] for i in indices]
            df_export = self.df[colunas_selecionadas]
            
            print(f"\n✅ {len(colunas_selecionadas)} colunas selecionadas:")
            print(df_export.head())
            
            nome = input("\n💾 Nome do ficheiro (sem extensão): ")
            df_export.to_excel(f"{nome}.xlsx", index=False)
            print(f"✅ Exportado como '{nome}.xlsx'")
            
        except Exception as e:
            print(f"❌ Erro: {e}")


def menu_principal():
    """Menu principal do programa"""
    analisador = AnalisadorExcel()
    
    while True:
        print("\n" + "="*60)
        print("📊 ANALISADOR DE FICHEIROS EXCEL")
        print("="*60)
        print("\n1. 📂 Carregar ficheiro Excel")
        print("2. ℹ️  Mostrar informação básica")
        print("3. 📈 Mostrar estatísticas")
        print("4. 🔍 Filtrar dados")
        print("5. 📊 Agrupar e resumir dados")
        print("6. 📋 Exportar colunas específicas")
        print("0. 🚪 Sair")
        
        escolha = input("\n➡️  Escolha uma opção: ")
        
        if escolha == "1":
            caminho = input("\n📁 Caminho do ficheiro Excel: ")
            analisador.carregar_ficheiro(caminho)
        
        elif escolha == "2":
            analisador.mostrar_info_basica()
        
        elif escolha == "3":
            analisador.mostrar_estatisticas()
        
        elif escolha == "4":
            analisador.filtrar_por_coluna()
        
        elif escolha == "5":
            analisador.agrupar_dados()
        
        elif escolha == "6":
            analisador.exportar_colunas_especificas()
        
        elif escolha == "0":
            print("\n👋 Até breve!")
            break
        
        else:
            print("\n❌ Opção inválida!")
        
        input("\nPressione ENTER para continuar...")


if __name__ == "__main__":
    print("""
    ╔═══════════════════════════════════════════════════════╗
    ║     ANALISADOR E FILTRADOR DE FICHEIROS EXCEL         ║
    ║                                                       ║
    ║  Ferramenta interativa para análise de dados Excel   ║
    ╚═══════════════════════════════════════════════════════╝
    """)
    
    # Verificar se pandas está instalado
    try:
        import pandas as pd
        menu_principal()
    except ImportError:
        print("❌ Erro: pandas não está instalado!")
        print("Execute: pip install pandas openpyxl")
