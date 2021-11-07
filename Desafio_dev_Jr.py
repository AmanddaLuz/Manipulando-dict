import pandas as pd
import os
from os import path
import win32com.client as win32
'''A referência para o último import é
pip install pywin32.'''

''' Menu de opções.'''
print('''      Acompanhamento do jogo!
      **************************
      MENU DE OPÇÕES:
      [1] para inserir
      [2] para consultar
      [3] para exportar em arquivo Excel
      [4] para encerrar ''')

'''Variáveis.'''
resumo = []
opcao = 0
maxi = 0
mini = 1000
qmini = 0
qmaxi = 0
cont = 0

'''Variável que identifica especificamente a pasta usuário 
do computador que executar o programa'''
pastaUsuario = path.join(path.expanduser('~'))

'''Variável que verifica se o arquivo existe no diretório do usuário.'''
exist = os.path.isfile(pastaUsuario + '\\resumo_dos_jos.xlsx')


'''Função utilizada para validar o input numérico, não aceitando espaços ou letras.
O input persiste até obter uma opção válida'''
def valida_num(msg):
    ok = False
    valor = 0
    while True:
        num = str(input(msg))
        if num.isnumeric():
            valor = int(num)
            ok = True
        else:
            print('Erro! Digite novamente!')
        if ok:
            break
    return valor


'''Função utilizada para validar o input nome, não aceitando espaços.
O input persiste até obter uma opção válida'''
def valida_nome(msg):
    while True:
        nome = input(msg)
        if nome == '':
            print('Erro! Digite um nome válido! ')
        else:
            break
    return nome


'''Validando o Menu de opções.'''
while True:
    opcao = valida_num('Digite a opção desejada de acordo com o Menu: ')
    if opcao == 1:
        cont += 1

        '''Recebe a pontuação marcada no jogo.'''
        pontuacao = valida_num('Digite os pontos do último jogo: ')
        if cont == 1:

            '''Se não existe nenhuma pontuação anterior,
            cria a lista com o jogo 1.'''
            lista = [1, pontuacao, pontuacao, pontuacao, 0, 0]
            mini = pontuacao
            maxi = pontuacao
            resumo.append(lista)
        else:

            '''Se já existe pontuação, compara e registra essas comparações
            em uma nova lista'''
            if pontuacao > maxi:
                maxi = pontuacao
                qmaxi += 1
            if pontuacao < mini:
                mini = pontuacao
                qmini += 1
            lista = [cont, pontuacao, mini, maxi, qmini, qmaxi]

            '''A variável resumo recebe o lançamento  e as comparações de cada jogo. '''
            resumo.append(lista)

        '''Gera a planilha a partir dos dados obtidos usando um dataframe e depois convertendo em excel')'''
        df = pd.DataFrame(resumo)
        df.rename(columns={0: 'Jogo',
                           1: 'Placar',
                           2: 'Mínimo da temporada',
                           3: 'Máximo da temporada',
                           4: 'Quebra de recorde min.',
                           5: 'Quebra de recorde máx'
                           }, inplace=True)

        '''Salva o documento na pasta usuário do respectivo usuário que executar o programa.'''
        df.to_excel(pastaUsuario + '\\resumo_dos_jogos.xlsx', index=False)

    elif opcao == 2:

        '''Verifica se o arquivo já existe, se sim, mostra o arquivo.'''
        if exist:
            df = pd.read_excel(pastaUsuario + '\\resumo_dos_jogos.xlsx')
            print(df)
        else:
            print('não existem dados para serem consultados! ')

    elif opcao == 3:

        '''Criando a integração com o outlook cadastrado
        para possível exportação do arquivo.'''
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)

        '''Variáveis para identificar os responsáveis pelo email'''
        jogador = valida_nome('Digite o nome do jogador: ')
        assinatura = valida_nome('Digite o nome do responsável pelo envio desse email: ')

        '''Variável para anexar um arquivo no email.'''
        anexo = pastaUsuario + '\\resumo_dos_jogos.xlsx'

        '''configurar as informações do seu e-mail CADASTRADO.'''
        email.To = "amandaomariano@hotmail.com"
        email.Subject = "E-mail automático - Resumo dos jogos."
        email.Attachments.Add(anexo)
        email.HTMLBody = f"""
                    <p>Olá, segue o resumo de desempenho dos jogos de {jogador}.</p>
                    <p>Abs,</p>
                    <p>Atenciosamente, {assinatura}</p>
                    """
        email.Send()
        print("Email Enviado")

    elif opcao == 4:
        print('Fim do programa.\n',
              'Obrigado(a). Até a próxima!')
        break
    else:
        print('Opção inválida!')

'''Eu gostaria de ter feito a validação do campo email
com um input para o email do remetente.

gostaria também de ter recuperado o arquivo salvo e adicionado novas informações
lançadas em outras datas, por exemplo.

Por conta do tempo, não consegui trabalhar nessas situações o 
suficiente para solucioná-las.

Mas ele cumpre os requisitos - ASSIM ESPERO! haha'''

