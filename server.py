# -*- coding: utf-8 -*-
"""
server.py
Servidor local e sincronizador automatico do GEOportal & BI-Obras BATTRE.
1. Sincroniza a planilha 'BI_PMO_Obras_BATTRE.xlsx' imediatamente ao iniciar.
2. Inicia o servidor web local (porta 8080 ou proxima disponivel).
3. Abre automaticamente o navegador padrao no GEOportal.
"""

import os
import sys
import time
import socket
import webbrowser
import http.server
import socketserver
from urllib.parse import urlparse

# Importa a sincronizacao
import sincronizar_dados

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

class GeoportalRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        # Desabilitar cache para garantir que sempre carregue dados frescos da planilha
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        # Rota de API para sincronizar sob demanda
        if parsed.path == '/api/sync':
            print("\n[SERVIDOR] Requisicao de sincronizacao recebida do navegador...")
            ok = sincronizar_dados.sincronizar()
            self.send_response(200 if ok else 500)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            resp = '{"status": "ok", "mensagem": "Planilha sincronizada com sucesso"}' if ok else '{"status": "erro"}'
            self.wfile.write(resp.encode('utf-8'))
            return

        # Redirecionar raiz para index.html
        if parsed.path == '/' or parsed.path == '':
            self.path = '/index.html'

        return super().do_GET()

    def do_POST(self):
        """Rota POST /api/save?file=nome.json — salva arquivo JSON na pasta do projeto."""
        from urllib.parse import urlparse, parse_qs
        import json as _json
        parsed = urlparse(self.path)
        if parsed.path == '/api/save':
            try:
                params   = parse_qs(parsed.query)
                filename = params.get('file', [None])[0]
                if not filename or '..' in filename or '/' in filename or '\\' in filename:
                    self.send_response(400)
                    self.end_headers()
                    self.wfile.write(b'{"status":"erro","msg":"Nome de arquivo invalido"}')
                    return
                length  = int(self.headers.get('Content-Length', 0))
                body    = self.rfile.read(length)
                # Validate JSON
                parsed_data = _json.loads(body)
                filepath = os.path.join(BASE_DIR, filename)
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(body.decode('utf-8'))
                print(f"[SERVIDOR] Arquivo salvo: {filename}")

                # Atualiza também o Excel em segundo plano se for plano ou estoque
                if filename in ('plano_acao.json', 'estoque.json'):
                    try:
                        sincronizar_dados.atualizar_excel_com_json(filename, parsed_data)
                        # Garante que o timestamp do JSON seja sempre mais recente ou igual ao do Excel
                        import time
                        time.sleep(0.05)
                        now = time.time()
                        os.utime(filepath, (now, now))
                    except Exception as err_xl:
                        print(f"[SERVIDOR] Aviso ao sincronizar Excel: {err_xl}")
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(('{"status":"ok","file":"' + filename + '"}').encode('utf-8'))
            except Exception as e:
                print(f"[SERVIDOR] Erro ao salvar arquivo: {e}")
                self.send_response(500)
                self.end_headers()
                self.wfile.write(('{"status":"erro","msg":"' + str(e) + '"}').encode('utf-8'))
            return
        self.send_response(405)
        self.end_headers()

def encontrar_porta_livre(porta_inicial=8080):
    porta = porta_inicial
    while porta < porta_inicial + 50:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', porta)) != 0:
                return porta
            porta += 1
    return 8080

def main():
    print("=" * 66)
    print("     BATTRE - GEOPORTAL & PAINEL EXECUTIVO PMO DE OBRAS")
    print("=" * 66)
    print(f"[1/3] Verificando e sincronizando planilha: BI_PMO_Obras_BATTRE.xlsx ...")
    
    try:
        sincronizar_dados.sincronizar()
    except Exception as e:
        print(f"[AVISO] Erro na sincronizacao inicial: {e}")

    porta = encontrar_porta_livre(8080)
    url = f"http://localhost:{porta}/index.html"

    print(f"\n[2/3] Iniciando servidor web local na porta {porta}...")
    
    # Permitir reuso do socket
    socketserver.TCPServer.allow_reuse_address = True
    
    try:
        with socketserver.TCPServer(('127.0.0.1', porta), GeoportalRequestHandler) as httpd:
            print(f"[3/3] Abrindo navegador em: {url}")
            print("\n" + "=" * 66)
            print("  SISTEMA PRONTO E MONITORANDO A PLANILHA!")
            print(f"  URL: {url}")
            print("  Toda vez que a planilha for salva no Excel, basta")
            print("  atualizar a pagina no navegador (F5).")
            print("  Mantenha esta janela aberta enquanto estiver utilizando o sistema.")
            print("=" * 66 + "\n")
            
            # Abre o navegador automaticamente
            webbrowser.open(url)
            
            # Executa o servidor
            httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[SERVIDOR] Encerrando GEOportal BATTRE...")
    except Exception as e:
        print(f"\n[ERRO] Falha ao iniciar servidor: {e}")
        input("Pressione Enter para sair...")

if __name__ == '__main__':
    main()
