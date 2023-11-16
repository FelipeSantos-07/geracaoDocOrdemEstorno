import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox
from openpyxl import load_workbook
from datetime import datetime

diretorio = ''

class VerificadorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de ordem de estorno")
        self.root.geometry("350x200")
        self.root.resizable(width=False, height=False)

        icon_path = "logoNeshop.ico"
        self.root.iconbitmap(icon_path)

        self.create_widgets()
    
    def select_directory(self):
        global diretorio
        diretorio = filedialog.askdirectory(parent=root, title='Diretório final da ordem de estorno')
        self.label.config(text=diretorio)

        if diretorio != "":
            self.verify_button.config(state=tk.NORMAL)
            self.verify_button.config(bg='blue', fg='white')
            # Altera o texto do botão
            self.btn_text.set("MUDAR O DIRETÓRIO")
        else:
            tkinter.messagebox.showinfo('ERRO', f'Nenhum diretório foi selecionado') 
    

    def create_widgets(self):
        label = tk.Label(root, text='SELECIONE O DIRETÓRIO')
        label.pack(pady=10)
        label.pack()

        # Variável de controle para o texto do botão
        self.btn_text = tk.StringVar()
        self.btn_text.set("ABRIR DIRETÓRIOS")

        button = tk.Button(root, textvariable=self.btn_text, command=self.select_directory)
        button.pack()

        self.label = tk.Label(self.root, text="")
        self.label.pack(pady=10)

        self.verify_button = tk.Button(self.root, text="GERAR ORDEM DE ESTORNO", command=self.verify_file, state=tk.DISABLED)
        self.verify_button.pack()


    def verify_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos de Texto", "*.txt")])
        if file_path:
            self.process_file(file_path)


    def process_file(self, file_path):
        global diretorio
        def mascaraCPFCNPJ(x):
            cpfcpnjCliente = x.replace(".", "").replace("-", "").replace("/", "")
            if len(cpfcpnjCliente) < 14 : # É CPF
                return f'{cpfcpnjCliente[:3]}.{cpfcpnjCliente[3:6]}.{cpfcpnjCliente[6:9]}-{cpfcpnjCliente[-2:]}'
            else: # É CNPJ
                return f'{cpfcpnjCliente[:2]}.{cpfcpnjCliente[2:5]}.{cpfcpnjCliente[5:8]}/{cpfcpnjCliente[8:12]}-{cpfcpnjCliente[-2:]}'
        
        def mascaraTelefone(numeroTelefone):
            if len(numeroTelefone) < 11:
                return f'+55 ({numeroTelefone[:2]}) {numeroTelefone[-8:-4]}-{numeroTelefone[-4:]}'
            else:
                return f'+55 ({numeroTelefone[:2]}) {numeroTelefone[-9:-4]}-{numeroTelefone[-4:]}'


        with open(file_path, "r", encoding="utf-8") as file:
            data = file.readlines()

        result = {}
        for line in data:
            # Verifica se a linha contém pelo menos dois elementos separados por ":"
            if ":" in line:
                title, value = map(str.strip, line.split(":", 1))
                result[title] = value
            else:
                # Trata a situação em que a linha não possui dois elementos
                #print(f"A linha não contém dois elementos separados por ':' - Ignorando: {line}")
                continue


        codigoOrdem = "-"
        numeroPedido = "-"
        propostaComercial = "-"
        nomeCliente = "-"
        telefoneContatoCliente = "-"
        cpfcpnjCliente = "-"
        emailCliente = "-"
        #tipoEstorno = "-"
        valorEstorno = "-"
        tipoPagamento = "-"
        tipoChavePix = "-"
        codigoChavePix = "-"
        nomeBancoCliente = "-"
        agenciaBancoCliente = "-"
        contaBancoCliente = "-"
        motivoDevolucao = "-"
        notaFiscalVenda = "-"
        nomeAtendente = "-"

        for title, value in result.items():
            #title = str(title)
            if value != "":
                if "Ordem" in title:  # Código da ordem
                    codigoOrdem = str(value)
                elif "pedido" in title:  # Número do pedido
                    numeroPedido = str(value)
                elif "comercial" in title:  # Proposta comercial
                    propostaComercial = str(value)
                elif "cliente" in title:  # Nome do cliente
                    nomeCliente = str(value)
                elif "contato" in title:  # Telefone contato
                    telefoneContatoCliente = str(value)
                elif "CPF/CNPJ" in title:  # CPF/CNPJ
                    cpfcpnjCliente = str(value)
                elif "Email" in title:  # Email
                    emailCliente = str(value)
                #elif "estorno" in title:  # Tipo do estorno
                #    tipoEstorno = str(value)
                elif "Valor" in title:  # Valor
                    valorEstorno = str(value)
                elif "pagamento" in title:  # Tipo pagamento
                    tipoPagamento = str(value)
                elif "Tipo" in title:  # Tipo Chave pix
                    tipoChavePix = str(value)
                elif "Chave" in title:  # Chave pix
                    codigoChavePix = str(value)
                elif "banco" in title:  # Nome banco
                    nomeBancoCliente = str(value)
                elif "Agência" in title:  # Agência
                    agenciaBancoCliente = str(value)
                elif "Conta" in title:  # Conta
                    contaBancoCliente = str(value)
                elif "devolução" in title:  # Motivo devolução
                    motivoDevolucao = str(value)
                elif "Nota" in title:  # Nota fiscal
                    notaFiscalVenda = str(value)
                elif "Atendente" in title:  # Atendente
                    nomeAtendente = str(value)
            #print(f"{title}: {value}")
        
        naoCriaOrdem = False
        contador = 0
        data_atual = datetime.now()
        
        if cpfcpnjCliente != "-":
            cpfcpnjCliente = mascaraCPFCNPJ(cpfcpnjCliente)
        
        if tipoChavePix == 'cpf' or 'CPF' or 'cnpj' or 'CNPJ':
            codigoChavePix = mascaraCPFCNPJ(codigoChavePix)
        
        for item in (codigoOrdem, numeroPedido, cpfcpnjCliente, emailCliente, valorEstorno, tipoPagamento, codigoChavePix, motivoDevolucao, notaFiscalVenda):
            itens = ['Código da ordem', 'Número do pedido', 'CPF/CNPJ', 'Email', 'Valor', 'Tipo de pagamento', 'Chave pix', 'Motivo da devolução', 'Nota fiscal']
            if item == '-':
                tkinter.messagebox.showinfo('ERRO', f'A ordem não pôde ser gerada por falta da(o): {itens[contador]}') 
                naoCriaOrdem = True 
                break
            contador += 1
            
        if not naoCriaOrdem:
            arquivoExcel = load_workbook(filename='OrdemEstorno.xlsx')
            edicaoDados = arquivoExcel.active
            edicaoDados['D8'] = data_atual.strftime("%d/%m/%Y")
            edicaoDados['J8'] = codigoOrdem
            edicaoDados['B8'] = numeroPedido
            edicaoDados['B9'] = propostaComercial
            edicaoDados['C16'] = nomeCliente.title()
            edicaoDados['A32'] = mascaraTelefone(telefoneContatoCliente)
            edicaoDados['C20'] = cpfcpnjCliente
            edicaoDados['F8'] = emailCliente
            #edicaoDados['E9'] = tipoEstorno
            edicaoDados['C8'] = float(valorEstorno)
            edicaoDados['C12'] = tipoPagamento
            edicaoDados['C15'] = tipoChavePix
            edicaoDados['C14'] = codigoChavePix
            edicaoDados['C17'] = nomeBancoCliente
            edicaoDados['C18'] = agenciaBancoCliente
            edicaoDados['C19'] = contaBancoCliente
            edicaoDados['A24'] = motivoDevolucao
            edicaoDados['A35'] = notaFiscalVenda
            edicaoDados['A29'] = nomeAtendente.upper()

            diretorio = str(diretorio).replace("/", "\\")
            arquivoExcel.save(filename=f'{diretorio}\\{codigoOrdem} - {numeroPedido}.xlsx')
            tkinter.messagebox.showinfo('SUCESSO!', f'A ordem nro. {codigoOrdem} foi gerada com sucesso!') 


if __name__ == "__main__":
    root = tk.Tk()
    app = VerificadorGUI(root)
    root.mainloop()


# ALTERAÇÕES: NÃO ACEITA A ',' SOMENTE O '.' -> VALOR