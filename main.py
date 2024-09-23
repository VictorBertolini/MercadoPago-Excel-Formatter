
from Functions import *

# @Author Victor Bertolini de Sousa

# The list of content of the lines that will be removed 
remove_lines_list = [
    "EXTRATO DE CONTA", 
    "51.453.100 JOAO VITOR SOUZA VILELA", 
    "CPF/CNPJ:", 
    "Periodo:", 
    "Saldo inicial:", 
    "Entradas:", 
    "Saidas:", 
    "DETALHE DOS MOVIMENTOS", 
    "Data Descrição ID da operação Valor Saldo",
    "Saldo final:",
    "Data de geração:",
    "Você tem alguma dúvida?",
    "ligue para",
    "Mercado Pago Instituição",
    "Encontre nossos canais"
    ]

# The operations that i want to attach and clean
key_operations_list = [
    "Liberação de dinheiro",
    "Pagamento com Código QR Pix",
    "Transferência Pix recebida",
    "Transferência Pix enviada"
    ]

get_bank_statement_to_excel_file("Extrato Agosto.pdf", remove_lines_list, key_operations_list)

print("Program finished")