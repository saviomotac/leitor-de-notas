import pandas as pd
from datetime import datetime
import re

class InterExcelReaderBuilder:
    def __init__(self, filepath):
        self.filepath = filepath
        self.parsed_data = {}

    def parse(self):
        df = pd.read_excel(self.filepath, sheet_name="Sheet1", header=None)

        header_index = df[df.apply(lambda row: row.astype(str).str.contains("ESPECIFICAÇÃO DO TÍTULO").any(), axis=1)].index[0]
        df.columns = df.iloc[header_index].fillna('').astype(str)
        df_clean = df.iloc[header_index + 1:].copy()
        df_clean = df_clean[df_clean['QUANTIDADE'].apply(lambda x: str(x).replace('.', '').isdigit())]

        self.parsed_data['negocios'] = []
        for _, row in df_clean.iterrows():
            titulo_completo = str(row['ESPECIFICAÇÃO DO TÍTULO']).strip()
            titulo = titulo_completo.split()[0] if titulo_completo else ""
            if "SUBTOTAL" in titulo.upper():
                continue
            self.parsed_data['negocios'].append({
                "titulo": titulo,
                "qtd": int(str(row['QUANTIDADE']).replace('.', '')),
                "preco": float(str(row['PREÇO DE LIQUIDAÇÃO(R$)']).replace(',', '.')),
                "valor_operacao": float(str(row['COMPRA/VENDA (R$)']).replace('.', '').replace(',', '.')),
                "operacao": str(row['C/V']).strip(),
                "obs": '' if str(row.get('OBS(*)', '')).lower() == 'nan' else str(row.get('OBS(*)', '')).strip()
            })

        self.agrupar_negociacoes()

        resumo_indices = df[df.iloc[:, 0] == 'RESUMO DOS NEGÓCIOS'].index.tolist()
        resumos_extraidos = []

        for idx in resumo_indices:
            bloco = df.iloc[idx+1:idx+10].copy().dropna(how='all')
            resumo = {}
            for _, row in bloco.iterrows():
                chave = str(row.iloc[0]).replace(":", "").strip().lower().replace(" ", "_")
                try:
                    valor = float(str(row.iloc[1]).replace('.', '').replace(',', '.'))
                    resumo[chave] = valor
                except:
                    continue
            resumos_extraidos.append(resumo)

        resumo_negocios = resumos_extraidos[-1] if resumos_extraidos else {}
        self.parsed_data['resumo_negocios'] = {
            k: v for k, v in resumo_negocios.items() if k != 'nan'
        }

        numero = next((row[i+1] for _, row in df.iterrows() for i, cell in enumerate(row[:-1])
                       if isinstance(cell, str) and "NUM NOTA" in cell.upper() and pd.notna(row[i+1])), 0)

        data_pregao = next((datetime.strptime(str(row[i+1]), "%d%m%Y") for _, row in df.iterrows() for i, cell in enumerate(row[:-1])
                            if isinstance(cell, str) and "DATA PREGÃO" in cell.upper() and pd.notna(row[i+1])), None)

        total_raw = next((row[i+1] for _, row in df.iterrows() for i, cell in enumerate(row[:-1])
                          if isinstance(cell, str) and "LIQ.(A+B) P/" in cell.upper() and pd.notna(row[i+1])), '0')

        total_str_cleaned = re.sub(r'[^0-9,.-]', '', str(total_raw)).replace('.', '').replace(',', '.')
        total = abs(float(total_str_cleaned)) if total_str_cleaned else 0.0

        self.parsed_data['numero'] = int(numero)
        self.parsed_data['data_pregao'] = data_pregao
        self.parsed_data['resumo_financeiro'] = {}
        self.parsed_data['total'] = total / 100

        self.calcular_custos()

        return self.parsed_data

    def agrupar_negociacoes(self):
        agrupados = {}
        for negocio in self.parsed_data['negocios']:
            titulo = negocio['titulo']
            operacao = negocio['operacao'].upper()
            if titulo.endswith('F'):
                titulo = titulo[:-1]
            chave = (titulo, operacao)
            if chave not in agrupados:
                agrupados[chave] = {
                    'titulo': titulo,
                    'qtd': 0,
                    'valor_operacao': 0.0,
                    'operacao': operacao,
                    'custo': 0.0,
                    'total_com_custo': 0.0
                }
            agrupados[chave]['qtd'] += negocio['qtd']
            agrupados[chave]['valor_operacao'] += negocio['valor_operacao']

        for ag in agrupados.values():
            if ag['qtd'] > 0:
                ag['preco'] = ag['valor_operacao'] / ag['qtd']
            else:
                ag['preco'] = 0.0

        self.parsed_data['negocios'] = list(agrupados.values())

    def calcular_custos(self):
        negocios = self.parsed_data['negocios']
        total_nota = self.parsed_data['total']

        soma_valores_absolutos = sum([abs(n['valor_operacao']) for n in negocios])
        soma_titulos_ajustado = sum([
            -n['valor_operacao'] if n['operacao'].upper() == 'V' else n['valor_operacao']
            for n in negocios
        ])

        for n in negocios:
            proporcao = abs(n['valor_operacao']) / soma_valores_absolutos if soma_valores_absolutos != 0 else 0
            custo = abs((total_nota - abs(soma_titulos_ajustado)) * proporcao)
            n['proporcao'] = proporcao
            n['custo'] = custo
            n['total_com_custo'] = n['valor_operacao'] + custo

    def imprimir_negocios(self):
        df = pd.DataFrame(self.parsed_data['negocios'])
        float_cols = df.select_dtypes(include=['float']).columns
        df[float_cols] = df[float_cols].applymap(lambda x: round(x, 2))
        colunas_ordenadas = ["titulo", "operacao", "qtd", "preco", "valor_operacao", "custo", "total_com_custo"]
        df = df[colunas_ordenadas]
        print("Data " + str(self.parsed_data['data_pregao'].strftime("%d/%m/%Y")))
        print(df.to_string(index=False))
        print("Total: R$ " + str(self.parsed_data['total']))