"""
Corrige a lista de obras no vínculo de folhas (aba Medido): separa a competência
da folha das demais e pede confirmação ao vincular em outro mês.

Rodar no VM, na pasta do sistema:  python aplicar_fix_obras.py
Guarda uma cópia do arquivo original antes de alterar.
"""
import shutil, sys
from datetime import datetime
from pathlib import Path

ARQ = Path(__file__).resolve().parent / "templates" / "index.html"
TROCAS = [('\n    // Obras disponíveis para dropdown (mesmo contrato, todos os períodos)\n    let obras=data.filter(d=>d.contrato_num===f.n_contrato);\n\n    const obraOpts=obras.map(d=>`<option value="${d.id}">${escQ(d.obra||\'\')}</option>`).join(\'\');\n    const selHtml=obras.length\n', '\n    // Obras do mesmo contrato. O Medido soma pela competência da obra, então as da\n    // competência da folha vêm primeiro; as de outros meses ficam separadas e com o mês no nome.\n    const obras=data.filter(d=>d.contrato_num===f.n_contrato);\n    const compFolha=(f.periodo||\'\').substring(0,7);\n    const porNome=(a,b)=>(a.obra||\'\').localeCompare(b.obra||\'\');\n    const doMes=obras.filter(d=>d.comp===compFolha).sort(porNome);\n    const outros=obras.filter(d=>d.comp!==compFolha).sort((a,b)=>(b.comp||\'\').localeCompare(a.comp||\'\')||porNome(a,b));\n    const repetidos=new Set(doMes.map(d=>d.obra).filter((o,i,arr)=>arr.indexOf(o)!==i));\n    const optMes=doMes.map(d=>`<option value="${d.id}">${escQ(d.obra||\'\')}${repetidos.has(d.obra)?` — prov. R$ ${fmt(d.provisao||0)}`:\'\'}</option>`).join(\'\');\n    const optOutros=outros.map(d=>`<option value="${d.id}" data-comp="${d.comp||\'\'}">${escQ(d.obra||\'\')} — ${compLabel(d.comp)}</option>`).join(\'\');\n    const selHtml=obras.length\n'), ('    const selHtml=obras.length\n      ?`<select id="rateio-sel-${key}" style="width:100%;font-size:11px;padding:4px 6px;margin-bottom:3px"><option value="">-- Selecionar obra --</option>${obraOpts}</select>`\n      :`<span style="font-size:11px;color:#ccc">Sem obras para ${f.n_contrato}</span>`;\n', '    const selHtml=obras.length\n      ?`<select id="rateio-sel-${key}" data-comp="${compFolha}" style="width:100%;font-size:11px;padding:4px 6px;margin-bottom:3px"><option value="">-- Selecionar obra --</option>`\n        +(doMes.length?`<optgroup label="${compLabel(compFolha)} (competência da folha)">${optMes}</optgroup>`:\'\')\n        +(outros.length?`<optgroup label="Outras competências — conta no mês da obra">${optOutros}</optgroup>`:\'\')\n        +`</select>`\n        +(doMes.length?\'\':`<div style="font-size:10px;color:#A32D2D;margin-bottom:3px">Nenhuma obra lançada em ${compLabel(compFolha)} para este contrato.</div>`)\n      :`<span style="font-size:11px;color:#ccc">Sem obras para ${f.n_contrato}</span>`;\n'), ("  if(valor<=0){toast('Informe um valor maior que zero.');return;}\n  await vincularInline(nFolha,obraId,valor,periodo);\n", "  if(valor<=0){toast('Informe um valor maior que zero.');return;}\n  const compObra=sel.selectedOptions[0]?.dataset.comp;\n  if(compObra&&!confirm(`Esta obra é da competência ${compLabel(compObra)}, mas a folha é de ${compLabel(sel.dataset.comp)}.\\n\\nO valor vai contar em ${compLabel(compObra)}. Continuar?`))return;\n  await vincularInline(nFolha,obraId,valor,periodo);\n")]

texto = ARQ.read_text(encoding="utf-8")
if all(novo in texto for _, novo in TROCAS):
    print("Correção já aplicada. Nada a fazer.")
    sys.exit(0)
faltando = [i + 1 for i, (antigo, _) in enumerate(TROCAS) if texto.count(antigo) != 1]
if faltando:
    print(f"O index.html deste computador é diferente do esperado (trecho(s) {faltando}).")
    print("Nada foi alterado. Mande o arquivo templates/index.html para o Claude ajustar.")
    sys.exit(1)
copia = ARQ.with_name(f"index.html.antes_fix_{datetime.now():%Y%m%d_%H%M%S}")
shutil.copy2(ARQ, copia)
for antigo, novo in TROCAS:
    texto = texto.replace(antigo, novo)
ARQ.write_text(texto, encoding="utf-8")
print(f"Correção aplicada. Cópia do original: {copia.name}")
print("Recarregue a página com Ctrl+F5 (se não mudar, reinicie o sistema).")
