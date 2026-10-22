"""Читаемые обозначения комментариев, сохраняемые кодовой страницей VBE."""
ALIASES = {'²': '^2', '³': '^3', '⁴': '^4', '×': 'x', '→': ' -> ', '←': ' <- ',
           '≤': '<=', '≥': '>=', 'μ': 'mu', 'λ': 'lambda', 'π': 'pi', 'σ': 'sigma',
           'ε': 'epsilon', 'κ': 'kappa', 'φ': 'phi', 'ψ': 'psi', 'β': 'beta',
           'Δ': 'Delta', 'Σ': 'sum', '∑': 'sum', '√': 'sqrt', '∞': 'inf',
           '−': '-', '⁻': '^-', '⁹': '9'}

def readable(text):
    for symbol, replacement in ALIASES.items(): text = text.replace(symbol, replacement)
    # Неподдерживаемый символ не должен молча стать вопросительным знаком.
    text.encode('cp1251', errors='strict')
    return text

def comments_only(source):
    result=[]
    for line in source.splitlines():
        quoted, index=False, 0
        while index < len(line):
            if line[index] == '"':
                if quoted and index+1 < len(line) and line[index+1] == '"': index+=2; continue
                quoted=not quoted
            elif line[index] == "'" and not quoted:
                line=line[:index]+readable(line[index:]); break
            index+=1
        result.append(line)
    return '\n'.join(result)+'\n'
