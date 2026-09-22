#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SUPRIMIR IMAGENS - remove todas as imagens de documentos PDF e DOCX.
Roda 100% offline (nenhuma conexao de rede e usada).

Uso:
    python suprimir_imagens.py                      -> abre janela para escolher arquivos
    python suprimir_imagens.py documento.pdf       -> processa um arquivo
    python suprimir_imagens.py pasta/               -> processa todos PDF/DOCX da pasta
    python suprimir_imagens.py documento.docx --manter-cabecalho
          (DOCX: preserva imagens de cabecalho/rodape, ex.: logotipo)

Saida (ao lado do original, que NUNCA e alterado):
    <nome>_SEM_IMAGENS.pdf / .docx
    <nome>_SEM_IMAGENS_log.txt  (hashes SHA-256, contagem e verificacao)

Dependencias (instalar uma vez):
    pip install pymupdf lxml
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


# ----------------------------------------------------------------- main
def processar(path, manter_cabecalho):
    ext = os.path.splitext(path)[1].lower()
    dst = caminho_saida(path)
    if ext == ".pdf":
        total, ok, detalhe = processar_pdf(path, dst)
    elif ext == ".docx":
        total, ok, detalhe = processar_docx(path, dst, manter_cabecalho)
    else:
        print(f"  ignorado (formato nao suportado): {path}")
        return None

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

    status = "OK" if ok else "ATENCAO: VERIFICACAO FALHOU - NAO USAR"
    print(f"  {os.path.basename(path)} -> {total} imagem(ns) suprimida(s) | {status}")
    if not ok:
        print(f"    {detalhe}")
    return ok


def main():
    ap = argparse.ArgumentParser(description="Remove imagens de documentos PDF/DOCX (offline).")
    ap.add_argument("entradas", nargs="*", help="arquivos ou pastas")
    ap.add_argument("--manter-cabecalho", action="store_true",
                    help="DOCX: preserva imagens de cabecalho/rodape (ex.: logotipo)")
    args = ap.parse_args()

    entradas = args.entradas
    if not entradas:
        try:
            import tkinter as tk
            from tkinter import filedialog
            tk.Tk().withdraw()
            entradas = list(filedialog.askopenfilenames(
                title="Selecione os documentos",
                filetypes=[("Documentos", "*.pdf *.docx")]))
        except Exception:
            ap.print_help()
            return
    arquivos = []
    for e in entradas:
        if os.path.isdir(e):
            for n in sorted(os.listdir(e)):
                if n.lower().endswith((".pdf", ".docx")) and "_SEM_IMAGENS" not in n:
                    arquivos.append(os.path.join(e, n))
        else:
            arquivos.append(e)

    falhas = 0
    for a in arquivos:
        try:
            if processar(a, args.manter_cabecalho) is False:
                falhas += 1
        except Exception as ex:
            falhas += 1
            print(f"  ERRO em {a}: {ex}")
    print(f"\nConcluido: {len(arquivos)} arquivo(s), {falhas} com problema.")
    if os.name == "nt" and not args.entradas:
        input("Enter para fechar...")


if __name__ == "__main__":
    main()
