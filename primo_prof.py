# lista = [2,]
# for num in range(3,100,2):
# 	for primo in lista:
# 		div = num / primo
# 		if num % primo == 0:
# 			break
# 		if div < primo:
# 			lista.append(num)
# 			break
# print(lista)

lista = [2, ]
# para cada número no intervalo de 3 até 200 (não incluindo o 200)
# com o passo 2, ou seja, tirando os números pares porque eles nunca serão primos (com excessão do dois).
for num in range(3, 200, 2):
    # para cada número primo dentro da minha lista, que começa somente com o 2.
    for primo in lista:
        # a divisão (que gerará um númeiro float) do num gerado pelo range pelos números primos
        # que estão na minha lista.
        # type(div) == float
        div = num / primo
        # se o resto da minha divisão inteira do meu num por um número primo for zero
        # então pare! pois o num não é um primo.
        if num % primo == 0:
            break
        # se a divisão (float) for menor que o número primo que está dividindo o num então pare!
        # o num é um número primo.
        if div < primo:
            # agora adicione esse número à minha lista de números.
            lista.append(num)
            break

print(lista)