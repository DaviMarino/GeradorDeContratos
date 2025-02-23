from docx import Document
import os
import comtypes.client
import pandas as pd
import logging

# Configuração do logging
log_filename = "processo.log"
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def substituir_variaveis_formatado(modelo_path, valores, output_docx):
    """
    Substitui variáveis em um documento Word mantendo a formatação.

    :param modelo_path: Caminho do modelo Word (.docx)
    :param valores: Dicionário de substituições {'var01': 'João', 'var02': '25 anos'}
    :param output_docx: Caminho do documento Word gerado
    """
    try:
        doc = Document(modelo_path)

        for paragrafo in doc.paragraphs:
            for run in paragrafo.runs:
                for chave, valor in valores.items():
                    if chave in run.text:
                        run.text = run.text.replace(f'{{{chave}}}', valor)

        doc.save(output_docx)
        logging.info(f"Arquivo Word gerado: {output_docx}")
    except Exception as e:
        logging.error(f"Erro ao substituir variáveis no documento {output_docx}: {e}")

def converter_para_pdf(input_docx, output_pdf):
    """
    Converte um arquivo Word para PDF usando o Microsoft Word (somente no Windows).

    :param input_docx: Caminho do arquivo Word (.docx)
    :param output_pdf: Caminho do arquivo PDF gerado
    """
    try:
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False  # Mantém o Word oculto

        doc = word.Documents.Open(os.path.abspath(input_docx))
        doc.SaveAs(os.path.abspath(output_pdf), FileFormat=17)  # 17 = PDF format
        doc.Close()
        word.Quit()

        logging.info(f"PDF gerado com sucesso: {output_pdf}")
    except Exception as e:
        logging.error(f"Erro ao converter {input_docx} para PDF: {e}")

def gerar_pdfs(modelo_path, lista_dados, pasta_saida, prefixo, nome):
    """
    Gera múltiplos PDFs a partir de um modelo Word mantendo a formatação.

    :param modelo_path: Caminho do modelo Word (.docx)
    :param lista_dados: Lista de dicionários {'var01': 'João', 'var02': '25 anos'}
    :param pasta_saida: Pasta onde os documentos gerados serão salvos
    """
    try:
        if prefixo == "": prefixo = "Documento"
        if nome == "": nome = "var01"

        if not os.path.exists(pasta_saida):
            os.makedirs(pasta_saida)

        for i, valores in enumerate(lista_dados):
            try:
                doc_name = lista_dados[i][nome].replace(" ", "_")  # Evita espaços no nome do arquivo
                output_docx = os.path.join(pasta_saida, f"{prefixo}_{doc_name}.docx")
                output_pdf = os.path.join(pasta_saida, f"{prefixo}_{doc_name}.pdf")

                substituir_variaveis_formatado(modelo_path, valores, output_docx)
                converter_para_pdf(output_docx, output_pdf)
                logging.info(f"Documento processado: {output_pdf}")

            except Exception as e:
                logging.error(f"Erro ao processar documento {i+1}: {e}")
    except Exception as e:
        logging.critical(f"Erro fatal no processo de geração de PDFs: {e}")

if __name__ == "__main__":
    try:
        modelo_excel = input("Caminho completo do arquivo EXCEL com as variáveis:")
        modelo_excel = modelo_excel if modelo_excel != "" else "C:\Git\CompressorPDF\modelo.xlsx"
        dados = pd.read_excel(modelo_excel).to_dict(orient='records')

        modelo_word = input("Caminho completo do arquivo WORD para usar como modelo:")
        modelo_word = modelo_word if modelo_word != "" else "C:\Git\CompressorPDF\modelo.docx"
        prefixo = input("Prefixo para o documento? ")
        nome = input("Gerar o nome do documento com a variável? ")

        pasta_saida = f"{prefixo}_criado"
        gerar_pdfs(modelo_word, dados, pasta_saida, prefixo, nome)

        logging.info("Processo finalizado com sucesso.")

    except Exception as e:
        logging.critical(f"Erro fatal no script principal: {e}")