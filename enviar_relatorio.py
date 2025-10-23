import os
import shutil
import papermill as pm
import win32com.client
from datetime import datetime

# Caminhos
notebook_path = "credito_modalidade.ipynb"
saida_dir = "saida"
modelo_nome = f"{datetime.today().strftime('%Y%m%d')}_MONITORAMENTO DE COMPONENTE.xlsx"
relatorio_path = os.path.join(saida_dir, modelo_nome)

# Pasta pública do OneDrive
destino_publico = r"C:\Users\Datasus\OneDrive - Ministério da Saúde\Coordenação de Gestão da Informação - Documentos\BOT'S\PUBLICO"
destino_final = os.path.join(destino_publico, modelo_nome)

def executar_notebook():
    print("🚀 Executando notebook...")
    pm.execute_notebook(notebook_path, "credito_modalidade.ipynb")
    print("✅ Notebook executado.")

def copiar_para_publico():
    if os.path.exists(relatorio_path):
        shutil.copy(relatorio_path, destino_final)
        print(f"📁 Relatório copiado para pasta pública:\n{destino_final}")
    else:
        print(f"❌ Relatório não encontrado em: {relatorio_path}")

def enviar_email():
    imagem_assinatura = os.path.abspath("img/assinatura.jpg")  # Caminho da imagem

    if not os.path.exists(relatorio_path):
        print(f"❌ Arquivo para envio não encontrado: {relatorio_path}")
        return

    outlook = win32com.client.Dispatch("Outlook.Application")
    email = outlook.CreateItem(0)

    email.To = "email"
    email.Subject = f"Relatório Diário - {datetime.today().strftime('%d/%m/%Y')}"

    # Corpo em HTML com imagem embutida
    email.HTMLBody = (
        "<p>Prezados,</p>"
        "<p>Segue em anexo o relatório diário gerado automaticamente.</p>"
        "<p>Atenciosamente,<br>Otavio Augusto - BOT</p>"
        '<img src="cid:assinatura_img">'
    )

    # Anexo do relatório
    email.Attachments.Add(os.path.abspath(relatorio_path))

    # Anexo da imagem com CID
    assinatura = email.Attachments.Add(imagem_assinatura)
    assinatura.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "assinatura_img")

    email.Send()
    print("📤 E-mail enviado com sucesso com imagem de assinatura.")


if __name__ == "__main__":
    executar_notebook()
    copiar_para_publico()
    enviar_email()
