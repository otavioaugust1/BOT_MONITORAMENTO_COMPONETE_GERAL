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
    try:
        pm.execute_notebook(notebook_path, notebook_path)
        print("✅ Notebook executado.")
    except Exception as e:
        print(f"❌ Erro ao executar notebook: {e}")

def copiar_para_publico():
    if os.path.exists(relatorio_path):
        try:
            shutil.copy(relatorio_path, destino_final)
            print(f"📁 Relatório copiado para pasta pública:\n{destino_final}")
        except PermissionError:
            print(f"⚠️ Permissão negada ao copiar o arquivo. Verifique se ele está aberto: {relatorio_path}")
    else:
        print(f"❌ Relatório não encontrado em: {relatorio_path}")

def enviar_email():
    imagem_assinatura = os.path.abspath("img/assinatura.jpg")

    if not os.path.exists(relatorio_path):
        print(f"❌ Arquivo para envio não encontrado: {relatorio_path}")
        return

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        email = outlook.CreateItem(0)

        email.To = "otavio.santos@saude.gov.br;gabriella.neves@saude.gov.br;theresa.nakagawa@saude.gov.br;felipe.cotrim@saude.gov.br;cginfo.ate@saude.gov.br;filipe.mauricio@saude.gov.br"
        email.Subject = f"Relatório Diário - {datetime.today().strftime('%d/%m/%Y')}"

        email.HTMLBody = (
            "<p>Prezados,</p>"
            "<p>Segue em anexo o relatório diário gerado automaticamente."
            "Temos novidades.. agora com CNES nas proposta que tinha apenas CNPJ... </p>"
            "<p>Atenciosamente,<br>Otavio Augusto - BOT</p>"
            '<img src="cid:assinatura_img">'
        )

        email.Attachments.Add(os.path.abspath(relatorio_path))

        assinatura = email.Attachments.Add(imagem_assinatura)
        assinatura.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "assinatura_img")

        email.Send()
        print("📤 E-mail enviado com sucesso com imagem de assinatura.")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")

def limpar_arquivos_em_uso(pasta):
    for arquivo in os.listdir(pasta):
        caminho_arquivo = os.path.join(pasta, arquivo)
        if os.path.isfile(caminho_arquivo):
            try:
                os.remove(caminho_arquivo)
                print(f"🗑️ Arquivo removido: {caminho_arquivo}")
            except PermissionError:
                print(f"⚠️ Arquivo em uso, não foi possível excluir: {caminho_arquivo}")

if __name__ == "__main__":
    limpar_arquivos_em_uso(r"C:\Users\Datasus\Downloads")  # Se necessário
    executar_notebook()
    copiar_para_publico()
    enviar_email()
