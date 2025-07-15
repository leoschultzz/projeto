def aplicar_censura(nome, censurar=True): # Censura todo o nome, exceto se censurar=False.
    if not censurar:
        return nome

    partes = nome.strip().split()
    if not partes:
        return nome

    primeiro_nome = partes[0]
    if len(partes) == 1:
        return primeiro_nome

    return primeiro_nome + ' ' + ' '.join(['*' * len(p) for p in partes[1:]])


def censurar_cpf(cpf, censurar=True):
    """
    Censura um CPF mantendo o formato: ***.***.***-**
    Funciona com ou sem pontos/traços no input.
    """
    if not censurar:
        return cpf

    # Remove caracteres não numéricos
    numeros = ''.join([c for c in cpf if c.isdigit()])

    if len(numeros) != 11:
        return cpf  # Não é um CPF válido

    # Retorna censurado no formato padrão
    return f'***.***.***-{numeros[-2:]}'
