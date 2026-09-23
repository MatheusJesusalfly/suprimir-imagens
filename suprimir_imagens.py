#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SUPRIMIR IMAGENS - remove todas as imagens de documentos PDF e DOCX.
Roda 100% offline (nenhuma conexao de rede e usada).

Uso:
    Dois cliques no SuprimirImagens.exe -> abre a janela
    Arrastar arquivos/pastas sobre o icone do .exe ou para dentro da janela
    Linha de comando (sem janela):
        python suprimir_imagens.py --cli documento.pdf pasta/ [--manter-cabecalho]

Saida (ao lado do original, que NUNCA e alterado):
    <nome>_SEM_IMAGENS.pdf / .docx
    <nome>_SEM_IMAGENS_log.txt  (hashes SHA-256, contagem e verificacao)

Dependencias (instalar uma vez):
    pip install pymupdf lxml customtkinter tkinterdnd2
"""

import argparse
import datetime
import hashlib
import io
import os
import re
import sys
import zipfile

PLACEHOLDER = "[IMAGEM SUPRIMIDA]"
MAGIC = {  # assinaturas de arquivos de imagem, para verificacao final
    b"\x89PNG": "PNG", b"\xff\xd8\xff": "JPEG", b"GIF8": "GIF",
    b"BM": "BMP", b"II*\x00": "TIFF", b"MM\x00*": "TIFF",
}
IMG_EXT = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp",
           ".emf", ".wmf", ".svg", ".jfif", ".heic", ".ico")


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for bloco in iter(lambda: f.read(1 << 20), b""):
            h.update(bloco)
    return h.hexdigest()


def caminho_saida(path):
    base, ext = os.path.splitext(path)
    return f"{base}_SEM_IMAGENS{ext}"


# ------------------------------------------------------------------ PDF
def processar_pdf(src, dst):
    import pymupdf

    doc = pymupdf.open(src)
    if doc.needs_pass:
        raise RuntimeError("PDF protegido por senha - remova a senha antes.")

    total = 0
    for page in doc:
        infos = page.get_image_info(xrefs=True)
        for info in infos:
            r = pymupdf.Rect(info["bbox"]) & page.rect
            if r.is_empty:
                r = page.rect  # imagem fora da area visivel: remove mesmo assim
            page.add_redact_annot(r, text=PLACEHOLDER if r.width > 80 else "",
                                  fontsize=8, fill=(0.85, 0.85, 0.85),
                                  text_color=(0.2, 0.2, 0.2), align=1)
        total += len(infos)
        # remove imagens preservando o texto que estiver sobre/ao lado delas
        page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_REMOVE,
                              graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                              text=pymupdf.PDF_REDACT_TEXT_NONE)
        # segunda passada: qualquer imagem referenciada que tenha sobrado
        for img in page.get_images(full=True):
            try:
                page.delete_image(img[0])
                total += 1
            except Exception:
                pass

    # anexos, miniaturas, XMP (pode conter thumbnail) e JavaScript
    doc.scrub(attached_files=True, embedded_files=True, thumbnails=True,
              xml_metadata=True, javascript=True, clean_pages=True,
              hidden_text=False, metadata=False, redactions=True,
              redact_images=0, remove_links=False, reset_fields=False,
              reset_responses=False)
    # varredura final: qualquer objeto de imagem que ainda exista no arquivo
    # (ex.: referenciado por recursos herdados ou compartilhados entre paginas)
    # tem o conteudo sobrescrito por 1 pixel branco. Assim nenhum dado de imagem
    # permanece no arquivo, mesmo que o objeto continue referenciado.
    for x in range(1, doc.xref_length()):
        try:
            if doc.xref_get_key(x, "Subtype")[1] != "/Image":
                continue
            doc.update_object(x, "<</Type/XObject/Subtype/Image/Width 1/Height 1"
                                 "/BitsPerComponent 8/ColorSpace/DeviceGray>>")
            doc.update_stream(x, b"\xff")
        except Exception:
            pass

    doc.save(dst, garbage=4, deflate=True, clean=True)
    doc.close()

    # verificacao independente no arquivo gerado
    chk = pymupdf.open(dst)
    restantes = 0
    for x in range(1, chk.xref_length()):
        try:
            if chk.xref_get_key(x, "Subtype")[1] == "/Image":
                # imagem "vazia" deixada pelo delete_image (1x1 transparente) nao conta
                w = int(chk.xref_get_key(x, "Width")[1] or 0)
                h = int(chk.xref_get_key(x, "Height")[1] or 0)
                if w > 1 or h > 1:
                    restantes += 1
        except Exception:
            pass
    inline = sum(len([i for i in p.get_image_info() if i.get("width", 0) > 1])
                 for p in chk)
    anexos = chk.embfile_count()
    chk.close()
    ok = restantes == 0 and inline == 0 and anexos == 0
    detalhe = (f"objetos de imagem restantes: {restantes} | "
               f"imagens renderizaveis: {inline} | anexos: {anexos}")
    return total, ok, detalhe


# ----------------------------------------------------------------- DOCX
def processar_docx(src, dst, manter_cabecalho=False):
    from lxml import etree

    NS = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "v": "urn:schemas-microsoft-com:vml",
        "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    }
    W = "{%s}" % NS["w"]
    REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    REL_REMOVER = ("/image", "/oleObject", "/package", "/thumbnail")

    zin = zipfile.ZipFile(src)
    nomes = zin.namelist()
    arquivos = {n: zin.read(n) for n in nomes}
    zin.close()

    def protegido(nome):
        base = os.path.basename(nome)
        return manter_cabecalho and (base.startswith("header") or base.startswith("footer"))

    total = 0
    # 1) substitui desenhos/objetos com imagem por texto marcador
    for nome in nomes:
        if not (nome.startswith("word/") and nome.endswith(".xml")) or protegido(nome):
            continue
        if "_rels" in nome:
            continue
        raiz = etree.fromstring(arquivos[nome])
        alvos = raiz.xpath(
            "//mc:AlternateContent[.//a:blip or .//v:imagedata or .//w:object]"
            " | //w:drawing[.//a:blip][not(ancestor::mc:AlternateContent)]"
            " | //w:pict[.//v:imagedata][not(ancestor::mc:AlternateContent)]"
            " | //w:object[not(ancestor::mc:AlternateContent)]",
            namespaces=NS)
        for el in alvos:
            pai = el.getparent()
            if pai is None:
                continue
            t = etree.Element(W + "t")
            t.text = PLACEHOLDER
            if pai.tag == W + "r":
                pai.replace(el, t)
            else:  # fora de um run: cria um run novo
                r = etree.Element(W + "r")
                r.append(t)
                pai.replace(el, r)
            total += 1
        arquivos[nome] = etree.tostring(raiz, xml_declaration=True,
                                        encoding="UTF-8", standalone=True)

    # 2) remove relacionamentos de imagem/OLE e as partes correspondentes
    partes_removidas = set()
    for nome in nomes:
        if not nome.endswith(".rels"):
            continue
        dono = nome.replace("_rels/", "").replace(".rels", "")
        if protegido(dono):
            continue
        raiz = etree.fromstring(arquivos[nome])
        pasta = os.path.dirname(os.path.dirname(nome))
        for rel in list(raiz):
            tipo = rel.get("Type", "")
            if any(tipo.endswith(t) for t in REL_REMOVER) and rel.get("TargetMode") != "External":
                alvo = os.path.normpath(os.path.join(pasta, rel.get("Target"))).replace("\\", "/").lstrip("/")
                if rel.get("Target", "").startswith("/"):
                    alvo = rel.get("Target").lstrip("/")
                partes_removidas.add(alvo)
                raiz.remove(rel)
        arquivos[nome] = etree.tostring(raiz, xml_declaration=True,
                                        encoding="UTF-8", standalone=True)

    # imagens orfas em media/ e embeddings/ tambem saem
    for n in nomes:
        if protegido_midia(n, arquivos, manter_cabecalho):
            continue
        if ("/media/" in n or "/embeddings/" in n or n.startswith("docProps/thumbnail")):
            partes_removidas.add(n)

    for p in partes_removidas:
        if not protegido_midia(p, arquivos, manter_cabecalho):
            arquivos.pop(p, None)

    # 3) limpa [Content_Types].xml
    ct = etree.fromstring(arquivos["[Content_Types].xml"])
    for ov in list(ct):
        pn = (ov.get("PartName") or "").lstrip("/")
        if pn and pn not in arquivos:
            ct.remove(ov)
    arquivos["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True,
                                                    encoding="UTF-8", standalone=True)

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in nomes:
            if n in arquivos:
                zout.writestr(n, arquivos[n])

    # verificacao: nenhum arquivo de imagem (por extensao ou assinatura) fora das excecoes
    restantes = []
    with zipfile.ZipFile(dst) as z:
        for n in z.namelist():
            if protegido_midia(n, arquivos, manter_cabecalho):
                continue
            dados = z.read(n)[:8]
            if n.lower().endswith(IMG_EXT) or any(dados.startswith(m) for m in MAGIC):
                restantes.append(n)
    ok = not restantes
    detalhe = "arquivos de imagem restantes: " + (", ".join(restantes) if restantes else "0")
    if manter_cabecalho:
        detalhe += " (imagens de cabecalho/rodape preservadas por opcao)"
    return total, ok, detalhe


def protegido_midia(nome, arquivos, manter_cabecalho):
    """Com --manter-cabecalho, preserva midias ainda referenciadas por header/footer
    (as referencias do corpo ja foram removidas antes desta checagem)."""
    if not manter_cabecalho or "/media/" not in nome:
        return False
    alvo = os.path.basename(nome)
    usado_cab = False
    for n, dados in arquivos.items():
        if not n.endswith(".rels") or not isinstance(dados, bytes):
            continue
        if ("media/" + alvo).encode() in dados:
            base = os.path.basename(n)
            if base.startswith("header") or base.startswith("footer"):
                usado_cab = True
    return usado_cab


# ----------------------------------------------------------------- processamento
def processar(path, manter_cabecalho):
    """Processa um arquivo e devolve (total, ok, detalhe, caminho_gerado)."""
    ext = os.path.splitext(path)[1].lower()
    dst = caminho_saida(path)
    if ext == ".pdf":
        total, ok, detalhe = processar_pdf(path, dst)
    elif ext == ".docx":
        total, ok, detalhe = processar_docx(path, dst, manter_cabecalho)
    else:
        raise RuntimeError("formato nao suportado")

    log = os.path.splitext(dst)[0] + "_log.txt"
    with open(log, "w", encoding="utf-8") as f:
        f.write("SUPRESSAO DE IMAGENS - REGISTRO\n")
        f.write(f"Data/hora:        {datetime.datetime.now():%d/%m/%Y %H:%M:%S}\n")
        f.write(f"Original:         {os.path.abspath(path)}\n")
        f.write(f"SHA-256 original: {sha256(path)}\n")
        f.write(f"Gerado:           {os.path.abspath(dst)}\n")
        f.write(f"SHA-256 gerado:   {sha256(dst)}\n")
        f.write(f"Imagens suprimidas: {total}\n")
        f.write(f"Verificacao:      {'OK' if ok else 'FALHOU'} - {detalhe}\n")
    return total, ok, detalhe, dst


def expandir(caminhos):
    arquivos = []
    for c in caminhos:
        if os.path.isdir(c):
            for n in sorted(os.listdir(c)):
                if n.lower().endswith((".pdf", ".docx")) and "_SEM_IMAGENS" not in n:
                    arquivos.append(os.path.join(c, n))
        elif c.lower().endswith((".pdf", ".docx")) and "_SEM_IMAGENS" not in c:
            arquivos.append(c)
    return arquivos


def abrir_pasta(pasta):
    import subprocess
    if os.name == "nt":
        os.startfile(pasta)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", pasta])
    else:
        subprocess.Popen(["xdg-open", pasta])


# ----------------------------------------------------------------- interface
def iniciar_interface(iniciais):
    import threading
    import tkinter as tk
    from tkinter import filedialog
    import customtkinter as ctk

    try:  # arrastar e soltar (opcional)
        from tkinterdnd2 import TkinterDnD, DND_FILES
        class Base(ctk.CTk, TkinterDnD.DnDWrapper):
            def __init__(self, *a, **k):
                super().__init__(*a, **k)
                self.TkdndVersion = TkinterDnD._require(self)
        DND = True
    except Exception:
        Base, DND = ctk.CTk, False

    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")

    AZUL, VERDE, VERMELHO, CINZA = "#1f6aa5", "#2e8b57", "#c0392b", ("gray45", "gray65")

    app = Base()
    app.title("Suprimir Imagens")
    app.geometry("700x640")
    app.minsize(600, 560)

    estado = {"arquivos": [], "linhas": {}, "rodando": False, "ultima_pasta": None}

    # cabecalho
    topo = ctk.CTkFrame(app, fg_color="transparent")
    topo.pack(fill="x", padx=24, pady=(20, 6))
    ctk.CTkLabel(topo, text="Suprimir Imagens", font=ctk.CTkFont(size=22, weight="bold")).pack(anchor="w")
    ctk.CTkLabel(topo, text="Remove as imagens de documentos PDF e Word. Funciona offline — nada sai do computador.",
                 text_color=CINZA, font=ctk.CTkFont(size=13)).pack(anchor="w")

    # area de soltar
    zona = ctk.CTkFrame(app, corner_radius=12, border_width=2, border_color=("gray75", "gray30"))
    zona.pack(fill="x", padx=24, pady=(14, 8))
    ctk.CTkLabel(zona, text="⬇", font=ctk.CTkFont(size=26), text_color=CINZA).pack(pady=(14, 0))
    ctk.CTkLabel(zona, text="Arraste arquivos ou pastas aqui" if DND else "Selecione os arquivos ou uma pasta",
                 font=ctk.CTkFont(size=14, weight="bold")).pack()
    ctk.CTkLabel(zona, text="PDF e DOCX", text_color=CINZA, font=ctk.CTkFont(size=12)).pack()
    bts = ctk.CTkFrame(zona, fg_color="transparent")
    bts.pack(pady=(8, 14))

    # rodape (empacotado de baixo para cima, para nunca ser cortado)
    rod = ctk.CTkFrame(app, fg_color="transparent")
    rod.pack(side="bottom", fill="x", padx=24, pady=(6, 20))
    bt_exec = ctk.CTkButton(rod, text="Suprimir imagens", width=170, height=38,
                            font=ctk.CTkFont(size=14, weight="bold"))
    bt_exec.pack(side="right")
    bt_pasta = ctk.CTkButton(rod, text="Abrir pasta", width=110, height=38, fg_color="transparent",
                             border_width=1, text_color=("gray10", "gray90"), state="disabled")
    bt_pasta.pack(side="right", padx=8)
    bt_limpar = ctk.CTkButton(rod, text="Limpar lista", width=100, height=38, fg_color="transparent",
                              border_width=1, text_color=("gray10", "gray90"))
    bt_limpar.pack(side="left")

    lbl_status = ctk.CTkLabel(app, text="", font=ctk.CTkFont(size=12), text_color=CINZA)
    lbl_status.pack(side="bottom", anchor="w", padx=26)

    barra = ctk.CTkProgressBar(app, height=6)
    barra.set(0)
    barra.pack(side="bottom", fill="x", padx=24, pady=(4, 2))

    chk_var = tk.BooleanVar(value=False)
    ctk.CTkCheckBox(app, text="Manter imagens de cabeçalho e rodapé (Word) — ex.: logotipo",
                    variable=chk_var, font=ctk.CTkFont(size=12)).pack(side="bottom", anchor="w", padx=26, pady=(4, 8))

    # lista
    cab_lista = ctk.CTkFrame(app, fg_color="transparent")
    cab_lista.pack(fill="x", padx=24, pady=(6, 2))
    lbl_qtd = ctk.CTkLabel(cab_lista, text="Nenhum arquivo selecionado", text_color=CINZA,
                           font=ctk.CTkFont(size=12))
    lbl_qtd.pack(side="left")
    lista = ctk.CTkScrollableFrame(app, corner_radius=10, height=140)
    lista.pack(fill="both", expand=True, padx=24, pady=(0, 4))

    def atualizar_qtd():
        n = len(estado["arquivos"])
        lbl_qtd.configure(text="Nenhum arquivo selecionado" if n == 0 else f"{n} arquivo(s)")

    def adicionar(caminhos):
        if estado["rodando"]:
            return
        for a in expandir(caminhos):
            if a in estado["linhas"]:
                continue
            linha = ctk.CTkFrame(lista, fg_color="transparent")
            linha.pack(fill="x", pady=2)
            icone = "PDF" if a.lower().endswith(".pdf") else "DOC"
            ctk.CTkLabel(linha, text=icone, width=40, corner_radius=6,
                         fg_color=("#fde2e1", "#5a2323") if icone == "PDF" else ("#dde8f7", "#1d3552"),
                         font=ctk.CTkFont(size=10, weight="bold")).pack(side="left", padx=(4, 8))
            ctk.CTkLabel(linha, text=os.path.basename(a), anchor="w").pack(side="left", fill="x", expand=True)
            st = ctk.CTkLabel(linha, text="Aguardando", text_color=CINZA, font=ctk.CTkFont(size=12))
            st.pack(side="right", padx=8)
            estado["arquivos"].append(a)
            estado["linhas"][a] = (linha, st)
        atualizar_qtd()

    def limpar():
        if estado["rodando"]:
            return
        for linha, _ in estado["linhas"].values():
            linha.destroy()
        estado["arquivos"].clear()
        estado["linhas"].clear()
        barra.set(0)
        lbl_status.configure(text="")
        bt_pasta.configure(state="disabled")
        atualizar_qtd()

    def status(a, texto, cor):
        app.after(0, lambda: estado["linhas"][a][1].configure(text=texto, text_color=cor))

    def executar():
        if estado["rodando"]:
            return
        pend = [a for a in estado["arquivos"]]
        if not pend:
            lbl_status.configure(text="Adicione ao menos um arquivo.", text_color=VERMELHO)
            return
        estado["rodando"] = True
        bt_exec.configure(state="disabled", text="Processando…")
        manter = chk_var.get()

        def trabalho():
            falhas = 0
            for i, a in enumerate(pend, 1):
                status(a, "Processando…", AZUL)
                try:
                    total, ok, detalhe, dst = processar(a, manter)
                    estado["ultima_pasta"] = os.path.dirname(os.path.abspath(dst))
                    if ok:
                        status(a, f"✓  {total} imagem(ns) removida(s)", VERDE)
                    else:
                        falhas += 1
                        status(a, "✗  Verificação falhou — não usar", VERMELHO)
                except Exception as ex:
                    falhas += 1
                    msg = str(ex)
                    status(a, "✗  " + (msg[:40] + "…" if len(msg) > 40 else msg), VERMELHO)
                app.after(0, lambda v=i / len(pend): barra.set(v))

            def fim():
                estado["rodando"] = False
                bt_exec.configure(state="normal", text="Suprimir imagens")
                bt_pasta.configure(state="normal")
                if falhas:
                    lbl_status.configure(text=f"Concluído com {falhas} problema(s).", text_color=VERMELHO)
                else:
                    lbl_status.configure(text="Concluído. Arquivos e logs gerados ao lado dos originais.",
                                         text_color=VERDE)
            app.after(0, fim)

        threading.Thread(target=trabalho, daemon=True).start()

    def sel_arquivos():
        fs = filedialog.askopenfilenames(title="Selecione os documentos",
                                         filetypes=[("Documentos", "*.pdf *.docx")])
        adicionar(list(fs))

    def sel_pasta():
        p = filedialog.askdirectory(title="Selecione a pasta")
        if p:
            adicionar([p])

    ctk.CTkButton(bts, text="Selecionar arquivos", width=150, command=sel_arquivos).pack(side="left", padx=4)
    ctk.CTkButton(bts, text="Selecionar pasta", width=130, fg_color="transparent", border_width=1,
                  text_color=("gray10", "gray90"), command=sel_pasta).pack(side="left", padx=4)
    bt_exec.configure(command=executar)
    bt_limpar.configure(command=limpar)
    bt_pasta.configure(command=lambda: estado["ultima_pasta"] and abrir_pasta(estado["ultima_pasta"]))

    if DND:
        def soltar(ev):
            adicionar(list(app.tk.splitlist(ev.data)))
        app.drop_target_register(DND_FILES)
        app.dnd_bind("<<Drop>>", soltar)

    adicionar(iniciais)
    app.mainloop()


def main():
    ap = argparse.ArgumentParser(description="Remove imagens de documentos PDF/DOCX (offline).")
    ap.add_argument("entradas", nargs="*", help="arquivos ou pastas")
    ap.add_argument("--manter-cabecalho", action="store_true",
                    help="DOCX: preserva imagens de cabecalho/rodape (ex.: logotipo)")
    ap.add_argument("--cli", action="store_true", help="processa sem abrir a janela")
    args = ap.parse_args()

    if args.cli:
        for a in expandir(args.entradas):
            try:
                total, ok, detalhe, _ = processar(a, args.manter_cabecalho)
                print(f"{os.path.basename(a)} -> {total} imagem(ns) | {'OK' if ok else 'FALHOU: ' + detalhe}")
            except Exception as ex:
                print(f"ERRO em {a}: {ex}")
        return
    iniciar_interface(args.entradas)


if __name__ == "__main__":
    main()
