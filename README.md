# Manipulando-dict
"""Arquivo, onde faço o tratamento de uma lista de dicionários filtrando alguns pré-requisitos:
Fazer um exercício que leia o dicionário e efetue as seguintes alteracoes:
1 - Troque o nome completo do Estado pela sua respectiva sigla
2 - Sempre que houver um dado desconhecido trocar pela variável nula
3 - O email: acaso esteja como desconhecido trocar pelo modelo "nome.sobrenome@gmail.com", garantir que TODOS
OS EMAIL estejam em letra minúscula (low case)
4 - Trocar o nome do curso "A melhor linguagem do mundo" para "Python"
5 - Transformar o valor de cursos para o seguinte dicionário:
'cursos': {'Quantidade de cursos':
'Aluno Aplicado': True or False
'Aluno da melhor professora': True or False
'cursos do aluno': [lista dos cursos do aluno]}
O Aluno será aplicado se fizer mais de 2 cursos
O Aluno será aluno da melhor professora de estiver no curso de Python
*********************************************
Um exemplo de resolução:
{'usuarios': [{'nome completo': 'Matheus Esteque',
'Estado': 'Roraima',
'email': 'Matheus.Esteque@gmail.com',
'cursos': ['Pascal', 'JavaScript', 'C', 'Assembly', 'C++', 'Rusty'],
'phone': 981972367,
'endereço': 'Estrada Girassol',
'CEP': 1549431}, ]}

resposta esperada:
[{'nome completo': 'Matheus Esteque',
'Estado': 'RR',
'email': 'matheus.esteque@gmail.com',
'cursos': {'Quantidade de cursos': 6
'Aluno Aplicado': True
'Aluno da melhor professora': False
'cursos do aluno': ['Pascal', 'JavaScript', 'C', 'Assembly', 'C++', 'Rusty']},
'phone': 981972367,
'endereço': 'Estrada Girassol',
'CEP': 1549431}]
"""
