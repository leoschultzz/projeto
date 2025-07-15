def executar(igpm):
    import pandas as pd
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    import os
    from datetime import datetime
    from collections import defaultdict

    # Pasta de saída
    os.makedirs("pdfs_formatados", exist_ok=True)

    # Carrega dados
    df = pd.read_excel("resultado_dados.xlsx")

    # Parâmetros financeiros
    data_inicial = datetime(2024, 7, 31)
    data_final = datetime(2025, 4, 22)
    dias = (data_final - data_inicial).days
    meses = dias / 30.44
    juros_mensal = 0.01
    multa = 0.10

    # Estilo padrão
    styles = getSampleStyleSheet()
    normal = styles["Normal"]

    # Lista para armazenar os dados atualizados
    dados_atualizados = []

    def gerar_pdf(nome_arquivo, nome_pessoa, valor_parcela):
        valor_parcela = float(str(valor_parcela).replace('.', '').replace(',', '.'))
        valor_corrigido = round(valor_parcela * (1 + igpm), 2)
        valor_juros = round(valor_corrigido * juros_mensal * meses, 2)
        total_atualizado = round(valor_corrigido + valor_juros, 2)
        valor_multa = round(total_atualizado * multa, 2)
        total = round(total_atualizado + valor_multa, 2)

        doc = SimpleDocTemplate(
            f"pdfs_formatados/{nome_arquivo}.pdf",
            pagesize=A4,
            rightMargin=30, leftMargin=30,
            topMargin=30, bottomMargin=30
        )
        elements = []

        elements.append(Paragraph(f"<b>Atualização das Parcelas de AFUBRA X {nome_pessoa.upper()}</b>", styles['Title']))
        elements.append(Spacer(1, 12))

        texto = (
            "Forma de Cálculo:<br/>"
            "Parcelas Atualizadas Individualmente<br/>"
            "De 31/07/2024 a 22/04/2025 p/ IGPM<br/>"
            "Pró-Rata Nominal no 1º mês e Pró-Rata Nominal no último mês<br/>"
            "IGPM = Índice Geral de Preços do Mercado (FGV)<br/><br/>"
            "Forma dos Juros:<br/>"
            "De 31/07/2024 a 22/04/2025 juros Legais de 1,00 % ao mês, sobre o valor<br/>"
            "corrigido, sem capitalização<br/><br/>"
            "Multa de 10,00 % sobre o valor corrigido + juros"
        )
        elements.append(Paragraph(texto, normal))
        elements.append(Spacer(1, 18))

        headers = [
            "DATA", "DESCRIÇÃO", "VALOR DA PARCELA", "CORREÇÃO",
            "VALOR CORRIGIDO", "VALOR DOS JUROS", "TOTAL ATUALIZADO",
            "MULTA (10%)", "TOTAL"
        ]
        row = [
            "31/07/2024", "DÍVIDA",
            f"R$ {valor_parcela:,.2f}".replace('.', '#').replace(',', '.').replace('#', ','),
            f"{igpm * 100:.5f} %",
            f"R$ {valor_corrigido:,.2f}".replace('.', '#').replace(',', '.').replace('#', ','),
            f"R$ {valor_juros:,.2f}".replace('.', '#').replace(',', '.').replace('#', ','),
            f"R$ {total_atualizado:,.2f}".replace('.', '#').replace(',', '.').replace('#', ','),
            f"R$ {valor_multa:,.2f}".replace('.', '#').replace(',', '.').replace('#', ','),
            f"R$ {total:,.2f}".replace('.', '#').replace(',', '.').replace('#', ',')
        ]

        data = [headers, row]

        colWidths = [50, 60, 75, 50, 70, 70, 75, 45, 40]

        table = Table(data, colWidths=colWidths, repeatRows=1)

        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, 0), 6.5),
            ('FONTSIZE', (0, 1), (-1, -1), 6.5),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('ALIGN', (2, 1), (-1, 1), 'RIGHT'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ]))

        elements.append(table)
        doc.build(elements)

        return total  # Retorna o valor final calculado


    # Controle de Arquivos
    contador_arquivos = defaultdict(int)

    # Loop principal
    for _, row in df.iterrows():
        nome_arquivo_base = os.path.splitext(row['Arquivo'])[0]  # Pega apenas o nome sem extensão
        nome = row['Nome']
        valor = row['Dívida']

        # Atualiza contador para o arquivo
        contador_arquivos[nome_arquivo_base] += 1
        contador = contador_arquivos[nome_arquivo_base]

        # Nome real do arquivo
        if contador == 1:
            nome_arquivo_final = nome_arquivo_base
        else:
            nome_arquivo_final = f"{nome_arquivo_base}_{contador}"

        # Gera PDF e pega o valor TOTAL
        valor_total = gerar_pdf(nome_arquivo_final, nome, valor)

        # Adiciona no novo Excel
        dados_atualizados.append({
            'Arquivo': f"{nome_arquivo_final}.pdf",
            'Nome': nome,
            'Dívida Atualizada': f"{valor_total:,.2f}".replace('.', '#').replace(',', '.').replace('#', ','),
            'Número da Dívida': contador
        })

    # Criar o novo Excel 'dados_atualizados.xlsx'
    df_atualizado = pd.DataFrame(dados_atualizados)
    df_atualizado.to_excel('dados_atualizados.xlsx', index=False)

    print("✅ PDFs gerados!")
    print("✅ Arquivo 'dados_atualizados.xlsx' gerado!")

if __name__ == "__main__":
    executar()