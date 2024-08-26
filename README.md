# Introduction 
O projeto "Automação NFs" foi criado a partir da necessidade de baixar e organizar arquivos de notas fiscais vindo de emails no formato XML. O projeto lê os emails, baixa os arquivos em uma pasta temporária, lê o campo CNPJ do emitente da NF, e manda para uma pasta correspondente a este CNPJ (se a pasta não existir, ela é criada)

# Getting Started
O código não necessita de muitas configurações para funcionar: 
1. Ajustar provedor de email, endereço de email e senha de aplicativo
2. Ajustar caminho para pasta raiz, onde serão criadas as pastas dos CNPJs e arquivos temporários/inválidos

# Contribute
Esse projeto foi feito a partir de uma necessidade real de um cliente, a arquitetura .NET foi selecionada somente por motivos de familiaridade, para que não levasse muito tempo.