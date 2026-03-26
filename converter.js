/**
 * Valida um CNPJ nos formatos tradicional (só números) ou alfanumérico.
 * @param {string} cnpj - O CNPJ a ser validado (pode conter pontos, barras, hífens).
 * @returns {boolean} - Retorna true se o CNPJ for válido, false caso contrário.
 */
function validaCNPJ(cnpj) {
    // 1. Limpeza: remove tudo que não é letra ou número e converte para maiúsculas
    const cnpjLimpo = cnpj.replace(/[^0-9A-Za-z]/g, '').toUpperCase();

    // 2. Validações iniciais: tamanho deve ser 14 e não pode ser uma sequência inválida conhecida
    if (cnpjLimpo.length !== 14) return false;
    if (/^(\w)\1{13}$/.test(cnpjLimpo)) return false;

    // 3. Função auxiliar para converter caractere em valor (ASCII - 48)
    const charToValue = (char) => char.charCodeAt(0) - 48;

    // 4. Função auxiliar para calcular um dígito verificador
    const calcularDigito = (base, pesoInicial) => {
        let soma = 0;
        let peso = pesoInicial;

        for (let i = 0; i < base.length; i++) {
            const valor = charToValue(base[i]);
            soma += valor * peso;
            
            // Lógica dos pesos: vai de 9 a 2, depois recomeça
            peso = (peso === 2) ? 9 : peso - 1;
        }

        const resto = soma % 11;
        return resto < 2 ? 0 : 11 - resto;
    };

    // 5. Calcula os dígitos verificadores esperados
    const base12 = cnpjLimpo.substring(0, 12);
    const dv1Calculado = calcularDigito(base12, 5);
    const base13 = base12 + dv1Calculado;
    const dv2Calculado = calcularDigito(base13, 6);

    // 6. Compara os dígitos calculados com os dígitos informados
    const dvInformado = cnpjLimpo.substring(12, 14);
    return dvInformado === `${dv1Calculado}${dv2Calculado}`;
}

// --- Exemplos de uso ---
console.log(validaCNPJ('12.ABC.345/01DE-35'));   // true (exemplo válido)
console.log(validaCNPJ('12.ABC.345/01DE-99'));   // false (dígitos errados)
console.log(validaCNPJ('11.222.333/0001-81'));   // true (CNPJ numérico tradicional)
console.log(validaCNPJ('12.ABC.345/01DE-3'));    // false (tamanho incorreto)