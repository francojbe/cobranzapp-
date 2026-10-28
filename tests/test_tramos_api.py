import urllib.request
import json

def test_module(mod_id):
    # Select module
    req_sel = urllib.request.Request('http://127.0.0.1:8000/api/module/select',
                                     data=json.dumps({'module_id': mod_id}).encode('utf-8'),
                                     headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req_sel) as resp:
        res_sel = json.loads(resp.read().decode('utf-8'))
    
    # Load sample if not loaded (POST)
    if not res_sel.get('loaded'):
        req_load = urllib.request.Request(f'http://127.0.0.1:8000/api/load-sample?module_id={mod_id}', data=b'')
        with urllib.request.urlopen(req_load) as resp:
            json.loads(resp.read().decode('utf-8'))

    # Process
    req_proc = urllib.request.Request('http://127.0.0.1:8000/api/process', data=b'{}', headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req_proc) as resp:
        data = json.loads(resp.read().decode('utf-8'))
    
    tramos = data.get('resumen_global_tramos', {}).get('tramos', [])
    tot = data.get('resumen_global_tramos', {}).get('total_cartera_usd', 0)
    print(f"\n================ MODULO: {mod_id.upper()} (Total: ${tot:,.2f}) ================")
    for t in tramos:
        pct = (t['total_monto'] / tot * 100) if tot > 0 else 0
        print(f"Tramo {t['id']:<7} | {t['etiqueta']:<12} | Cuotas: {t['cuotas']:>5} | Clientes: {t['clientes']:>5} | Monto: ${t['total_monto']:>10.2f} | Pct Mora: {pct:>5.1f}%")

if __name__ == '__main__':
    test_module('general')
    test_module('avila')
