import json
from collections import defaultdict

data = json.load(open('historial.json', encoding='utf-8'))
movs = [m for m in data['movimientos'] if m['sku'] == 'CM']

por_mes = defaultdict(lambda: defaultdict(float))
for m in movs:
    mes = m['fecha'][:7]
    if m['tipo'] == 'venta':
        por_mes[mes][m['oficina']] += m['cantidad']

print('Ventas Carne Mechada por mes:')
for mes in sorted(por_mes):
    print(f"  {mes}: VIT={por_mes[mes]['VIT']:.0f}  PAT={por_mes[mes]['PAT']:.0f}")
