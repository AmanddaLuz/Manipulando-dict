primos = []
compostos = []

num = int(input('Digite, números primos até: '))
for dividendo in range(1, num+1):
    qtde_divisores = 0
    for divisor in range(1, dividendo+1):
        if dividendo % divisor == 0:
            qtde_divisores += 1
    if qtde_divisores == 2:
        primos.append(dividendo)
    if qtde_divisores > 2:
        compostos.append(dividendo)
print('\nSão {} nºs primos até o número {} escolhido. \nE são eles: {}'.format(len(primos), num, primos) + '\n'
      '\nSão {} nºs compostos até o {} escolhido. \nOs compostos são: {}'.format(len(compostos), num, compostos) + '\n'
      '\nO número 1 não é considerado nem primo nem composto.')
