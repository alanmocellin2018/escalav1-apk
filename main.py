import flet as ft
import json
import os
import tempfile
import calendar
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# DADOS FIXOS E CONFIGURAÇÕES
# ==============================================================================
CHAVE_STORAGE_DADOS = "escala_hospitalar_registros_v1"
CHAVE_STORAGE_EQUIPE = "escala_hospitalar_plantonistas_v1"

PLANTONISTAS_PADRAO = sorted([
    "Adriene Gomes de Moraes",
    "Caroline da Conceição Silva",
    "Fabiano Ribeiro de Castro",
    "Lara Thoany Alves de Oliveira Silva",
    "Letícia Ventura de Aquino Ferraz",
    "Marcela Mariani Lopes Soares Mocellin",
    "Silvio Zuliani Abdo Rodrigues",
    "Tânia Ramos Ribeiro",
    "Waldelysa Oliveira Scarpa"
])

TEXTO_NENHUM_EXTRA = "--- Nenhum (Apenas 1 plantonista) ---"

MESES = [
    "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
    "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
]

DIAS_SEMANA_PT = [
    "Segunda-feira", "Terça-feira", "Quarta-feira",
    "Quinta-feira", "Sexta-feira", "Sábado", "Domingo"
]

VALORES_PLANTAO = {
    "12h Semana": 1800.00,
    "6h Semana": 900.00,
    "12h FDS": 2000.00,
    "6h FDS": 1000.00
}


def main(page: ft.Page):
    page.title = "Escala Hospitalar"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.bgcolor = "#F1F5F9"
    page.padding = 12
    page.scroll = ft.ScrollMode.ADAPTIVE

    # ==========================================================================
    # ARMAZENAMENTO COMPATÍVEL COM APK ANDROID (CLIENT_STORAGE)
    # ==========================================================================
    def carregar_plantonistas():
        try:
            salvos = page.client_storage.get(CHAVE_STORAGE_EQUIPE)
            if salvos:
                lista = json.loads(salvos) if isinstance(salvos, str) else salvos
                if isinstance(lista, list) and len(lista) > 0:
                    return sorted(lista, key=lambda s: s.upper())
        except Exception:
            pass
        return list(PLANTONISTAS_PADRAO)

    def salvar_plantonistas():
        plantonistas.sort(key=lambda s: s.upper())
        try:
            page.client_storage.set(CHAVE_STORAGE_EQUIPE, json.dumps(plantonistas, ensure_ascii=False))
        except Exception:
            pass

    def carregar_dados():
        try:
            salvos = page.client_storage.get(CHAVE_STORAGE_DADOS)
            if salvos:
                lista = json.loads(salvos) if isinstance(salvos, str) else salvos
                if isinstance(lista, list):
                    return lista
        except Exception:
            pass
        return []

    def salvar_dados():
        try:
            page.client_storage.set(CHAVE_STORAGE_DADOS, json.dumps(registros, ensure_ascii=False))
        except Exception:
            pass

    plantonistas = carregar_plantonistas()
    registros = carregar_dados()
    estado = {"dia_atual": 1, " bytes_excel_temp": None}

    def formatar_moeda(valor):
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def mostrar_aviso(texto, cor="#1E3A8A"):
        page.snack_bar = ft.SnackBar(
            content=ft.Text(texto, color="white", weight=ft.FontWeight.BOLD),
            bgcolor=cor
        )
        page.snack_bar.open = True
        page.update()

    # ==========================================================================
    # SELETOR DE SALVAMENTO DE ARQUIVO NO ANDROID / PC
    # ==========================================================================
    def ao_escolher_local_salvar(e: ft.FilePickerResultEvent):
        if not e.path or not estado.get("caminho_temp_excel"):
            return
        try:
            with open(estado["caminho_temp_excel"], "rb") as f_orig:
                conteudo = f_orig.read()
            with open(e.path, "wb") as f_dest:
                f_dest.write(conteudo)
            mostrar_aviso(f"Planilha salva com sucesso!", "#16A34A")
        except Exception as err:
            mostrar_aviso(f"Erro ao salvar arquivo: {err}", "#DC2626")

    seletor_arquivo = ft.FilePicker(on_result=ao_escolher_local_salvar)
    page.overlay.append(seletor_arquivo)

    # ==========================================================================
    # ELEMENTOS VISUAIS DA INTERFACE
    # ==========================================================================
    hoje = datetime.now()
    dd_mes = ft.Dropdown(
        label="Mês",
        value=MESES[hoje.month - 1],
        options=[ft.dropdown.Option(m) for m in MESES],
        expand=True,
        bgcolor="white"
    )
    dd_ano = ft.Dropdown(
        label="Ano",
        value=str(hoje.year),
        options=[ft.dropdown.Option(str(a)) for a in range(2024, 2032)],
        width=115,
        bgcolor="white"
    )

    txt_dia_atual = ft.Text("Dia atual: 1", size=18, weight=ft.FontWeight.BOLD, color="#1E3A8A")
    txt_dia_semana = ft.Text("", size=13, weight=ft.FontWeight.BOLD, color="#2563EB")

    # Turno MANHÃ
    dd_manha = ft.Dropdown(
        label="Plantonista MANHÃ (07h - 19h)",
        value=plantonistas[0] if plantonistas else None,
        options=[ft.dropdown.Option(p) for p in plantonistas],
        bgcolor="white"
    )
    dd_manha_2 = ft.Dropdown(
        label="↳ 2º Plantonista Manhã (Divisão 6h)",
        value=TEXTO_NENHUM_EXTRA,
        options=[ft.dropdown.Option(TEXTO_NENHUM_EXTRA)] + [ft.dropdown.Option(p) for p in plantonistas],
        bgcolor="white",
        visible=False
    )

    chks_manha = {}
    for chave, val in VALORES_PLANTAO.items():
        chks_manha[chave] = ft.Checkbox(
            label=f"{chave} (R${int(val)})",
            value=(chave == "12h Semana"),
            on_change=lambda e, c=chave: selecionar_modalidade("manha", c)
        )

    # Turno NOITE
    dd_noite = ft.Dropdown(
        label="Plantonista NOITE (19h - 07h)",
        value=plantonistas[0] if plantonistas else None,
        options=[ft.dropdown.Option(p) for p in plantonistas],
        bgcolor="white"
    )
    dd_noite_2 = ft.Dropdown(
        label="↳ 2º Plantonista Noite (Divisão 6h)",
        value=TEXTO_NENHUM_EXTRA,
        options=[ft.dropdown.Option(TEXTO_NENHUM_EXTRA)] + [ft.dropdown.Option(p) for p in plantonistas],
        bgcolor="white",
        visible=False
    )

    chks_noite = {}
    for chave, val in VALORES_PLANTAO.items():
        chks_noite[chave] = ft.Checkbox(
            label=f"{chave} (R${int(val)})",
            value=(chave == "12h Semana"),
            on_change=lambda e, c=chave: selecionar_modalidade("noite", c)
        )

    txt_resumo_mes = ft.Text("", size=13, weight=ft.FontWeight.BOLD, color="#0F172A")
    coluna_historico = ft.Column(spacing=6)

    # ==========================================================================
    # LÓGICA DE NAVEGAÇÃO DE DIAS E BLOQUEIO SEMANA X FDS
    # ==========================================================================
    def obter_max_dias_mes():
        mes_idx = MESES.index(dd_mes.value) + 1
        ano = int(dd_ano.value)
        return calendar.monthrange(ano, mes_idx)[1]

    def obter_modalidade_marcada(mapa_chk):
        for chave, chk in mapa_chk.items():
            if chk.value:
                return chave
        return None

    def selecionar_modalidade(periodo, chave_escolhida):
        mapa = chks_manha if periodo == "manha" else chks_noite
        for chave, chk in mapa.items():
            chk.value = (chave == chave_escolhida)

        if periodo == "manha":
            dd_manha_2.visible = ("6h" in chave_escolhida)
            if not dd_manha_2.visible:
                dd_manha_2.value = TEXTO_NENHUM_EXTRA
        else:
            dd_noite_2.visible = ("6h" in chave_escolhida)
            if not dd_noite_2.visible:
                dd_noite_2.value = TEXTO_NENHUM_EXTRA
        page.update()

    def atualizar_info_dia():
        max_dias = obter_max_dias_mes()
        if estado["dia_atual"] > max_dias:
            estado["dia_atual"] = max_dias

        mes_idx = MESES.index(dd_mes.value) + 1
        ano = int(dd_ano.value)
        data_obj = datetime(ano, mes_idx, estado["dia_atual"])
        dia_sem_nome = DIAS_SEMANA_PT[data_obj.weekday()]
        eh_fds = data_obj.weekday() in (5, 6)

        txt_dia_atual.value = f"Dia atual: {estado['dia_atual']}"
        tipo_txt = "FINAL DE SEMANA" if eh_fds else "DIA DE SEMANA"
        txt_dia_semana.value = f"{estado['dia_atual']:02d}/{mes_idx:02d}/{ano} — {dia_sem_nome} ({tipo_txt})"
        txt_dia_semana.color = "#DC2626" if eh_fds else "#2563EB"

        # Bloqueia campos de FDS na semana e bloqueia Semana no FDS
        for chave in VALORES_PLANTAO.keys():
            eh_chave_fds = "FDS" in chave
            desabilitar = (not eh_chave_fds) if eh_fds else eh_chave_fds
            chks_manha[chave].disabled = desabilitar
            chks_noite[chave].disabled = desabilitar

        mod_m = obter_modalidade_marcada(chks_manha)
        eh_6h_m = "6h" in mod_m if mod_m else False
        selecionar_modalidade("manha", ("6h FDS" if eh_6h_m else "12h FDS") if eh_fds else ("6h Semana" if eh_6h_m else "12h Semana"))

        mod_n = obter_modalidade_marcada(chks_noite)
        eh_6h_n = "6h" in mod_n if mod_n else False
        selecionar_modalidade("noite", ("6h FDS" if eh_6h_n else "12h FDS") if eh_fds else ("6h Semana" if eh_6h_n else "12h Semana"))

    def ao_mudar_mes_ano(e=None):
        mes_nome = dd_mes.value
        ano = int(dd_ano.value)
        dias_lancados = [r["dia_num"] for r in registros if r["mes"] == mes_nome and r["ano"] == ano]
        max_dias = obter_max_dias_mes()
        if dias_lancados:
            estado["dia_atual"] = min(max(dias_lancados) + 1, max_dias)
        else:
            estado["dia_atual"] = 1
        atualizar_info_dia()
        atualizar_lista_historico()

    dd_mes.on_change = ao_mudar_mes_ano
    dd_ano.on_change = ao_mudar_mes_ano

    def voltar_dia(e):
        if estado["dia_atual"] > 1:
            estado["dia_atual"] -= 1
            atualizar_info_dia()

    def avancar_dia(e):
        if estado["dia_atual"] < obter_max_dias_mes():
            estado["dia_atual"] += 1
            atualizar_info_dia()

    # ==========================================================================
    # SALVAR DIA E ATUALIZAR HISTÓRICO
    # ==========================================================================
    def remover_registro_individual(idx_real):
        if 0 <= idx_real < len(registros):
            registros.pop(idx_real)
            salvar_dados()
            atualizar_lista_historico()
            mostrar_aviso("Registro removido.", "#DC2626")

    def atualizar_lista_historico():
        coluna_historico.controls.clear()
        mes_filtro = dd_mes.value
        ano_filtro = int(dd_ano.value)
        contador = 0

        for idx, reg in enumerate(registros):
            if reg["mes"] == mes_filtro and reg["ano"] == ano_filtro:
                contador += 1
                eh_dia = reg["noite_dia"].startswith("Dia")
                valor_exibido = formatar_moeda(reg["valor_dia"] if eh_dia else reg["valor_noite"])

                card_item = ft.Container(
                    bgcolor="white",
                    padding=10,
                    border_radius=8,
                    border=ft.border.all(1, "#CBD5E1"),
                    content=ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        controls=[
                            ft.Column(
                                spacing=2,
                                expand=True,
                                controls=[
                                    ft.Text(f"{reg['dia_str']} • {reg['nome']}", weight=ft.FontWeight.BOLD, size=13, color="#0F172A"),
                                    ft.Text(f"{reg['noite_dia']}  |  {reg['modalidade']}  |  {valor_exibido}", size=12, color="#475569")
                                ]
                            ),
                            ft.IconButton(
                                icon=ft.icons.DELETE_OUTLINE,
                                icon_color="#DC2626",
                                on_click=lambda e, i=idx: remover_registro_individual(i)
                            )
                        ]
                    )
                )
                coluna_historico.controls.append(card_item)

        txt_resumo_mes.value = f"Mês {mes_filtro}/{ano_filtro}: {contador} plantões"
        page.update()

    def salvar_dia_click(e):
        if not plantonistas:
            mostrar_aviso("Inclua pelo menos um plantonista antes de salvar!", "#DC2626")
            return

        mod_manha = obter_modalidade_marcada(chks_manha)
        mod_noite = obter_modalidade_marcada(chks_noite)

        mes_nome = dd_mes.value
        mes_idx = MESES.index(mes_nome) + 1
        ano = int(dd_ano.value)
        dia_num = estado["dia_atual"]
        max_dias = obter_max_dias_mes()

        data_obj = datetime(ano, mes_idx, dia_num)
        dia_semana = DIAS_SEMANA_PT[data_obj.weekday()]
        dia_formatado = f"{dia_num:02d} ({dia_semana[:3]})"

        nonlocal registros
        registros = [
            r for r in registros
            if not (r["dia_num"] == dia_num and r["mes"] == mes_nome and r["ano"] == ano)
        ]

        # 1. Manhã (Dia)
        val_m = VALORES_PLANTAO[mod_manha]
        registros.append({
            "nome": dd_manha.value,
            "dia_num": dia_num,
            "dia_str": dia_formatado,
            "dia_semana": dia_semana,
            "mes": mes_nome,
            "mes_num": mes_idx,
            "ano": ano,
            "noite_dia": "Dia (07h - 19h)",
            "modalidade": mod_manha,
            "valor_dia": val_m,
            "valor_noite": 0.0
        })

        if "6h" in mod_manha and dd_manha_2.value != TEXTO_NENHUM_EXTRA:
            registros.append({
                "nome": dd_manha_2.value,
                "dia_num": dia_num,
                "dia_str": dia_formatado,
                "dia_semana": dia_semana,
                "mes": mes_nome,
                "mes_num": mes_idx,
                "ano": ano,
                "noite_dia": "Dia (07h - 19h)",
                "modalidade": mod_manha,
                "valor_dia": val_m,
                "valor_noite": 0.0
            })

        # 2. Noite
        val_n = VALORES_PLANTAO[mod_noite]
        registros.append({
            "nome": dd_noite.value,
            "dia_num": dia_num,
            "dia_str": dia_formatado,
            "dia_semana": dia_semana,
            "mes": mes_nome,
            "mes_num": mes_idx,
            "ano": ano,
            "noite_dia": "Noite (19h - 07h)",
            "modalidade": mod_noite,
            "valor_dia": 0.0,
            "valor_noite": val_n
        })

        if "6h" in mod_noite and dd_noite_2.value != TEXTO_NENHUM_EXTRA:
            registros.append({
                "nome": dd_noite_2.value,
                "dia_num": dia_num,
                "dia_str": dia_formatado,
                "dia_semana": dia_semana,
                "mes": mes_nome,
                "mes_num": mes_idx,
                "ano": ano,
                "noite_dia": "Noite (19h - 07h)",
                "modalidade": mod_noite,
                "valor_dia": 0.0,
                "valor_noite": val_n
            })

        registros.sort(key=lambda r: (r["ano"], r["mes_num"], r["dia_num"], r["noite_dia"], r["nome"]))
        salvar_dados()
        atualizar_lista_historico()

        if estado["dia_atual"] < max_dias:
            estado["dia_atual"] += 1
            atualizar_info_dia()
            mostrar_aviso(f"Dia {dia_num:02d} salvo! Avançando para Dia {estado['dia_atual']:02d}.", "#16A34A")
        else:
            mostrar_aviso(f"Dia {dia_num:02d} salvo! Mês de {mes_nome} concluído!", "#16A34A")

    # ==========================================================================
    # INCLUIR / EXCLUIR PLANTONISTAS
    # ==========================================================================
    def atualizar_dropdowns_equipe():
        plantonistas.sort(key=lambda s: s.upper())
        opcoes_principais = [ft.dropdown.Option(p) for p in plantonistas]
        opcoes_extras = [ft.dropdown.Option(TEXTO_NENHUM_EXTRA)] + opcoes_principais

        dd_manha.options = opcoes_principais
        dd_noite.options = opcoes_principais
        dd_manha_2.options = opcoes_extras
        dd_noite_2.options = opcoes_extras

        if plantonistas:
            if dd_manha.value not in plantonistas:
                dd_manha.value = plantonistas[0]
            if dd_noite.value not in plantonistas:
                dd_noite.value = plantonistas[0]
        page.update()

    campo_novo_nome = ft.TextField(label="Nome completo do plantonista", autofocus=True)

    def confirmar_inclusao(e):
        nome = " ".join((campo_novo_nome.value or "").strip().split())
        if not nome:
            return
        if any(p.lower() == nome.lower() for p in plantonistas):
            mostrar_aviso("Esse plantonista já está cadastrado!", "#DC2626")
            return
        plantonistas.append(nome)
        salvar_plantonistas()
        atualizar_dropdowns_equipe()
        campo_novo_nome.value = ""
        modal_incluir.open = False
        mostrar_aviso(f"'{nome}' incluído em ordem alfabética!", "#16A34A")

    modal_incluir = ft.AlertDialog(
        title=ft.Text("➕ Incluir Plantonista"),
        content=campo_novo_nome,
        actions=[
            ft.TextButton("Cancelar", on_click=lambda e: setattr(modal_incluir, "open", False) or page.update()),
            ft.ElevatedButton("Salvar", bgcolor="#16A34A", color="white", on_click=confirmar_inclusao)
        ]
    )

    dd_excluir_med = ft.Dropdown(label="Selecione para remover")

    def confirmar_exclusao(e):
        nome_rem = dd_excluir_med.value
        if nome_rem and nome_rem in plantonistas:
            plantonistas.remove(nome_rem)
            salvar_plantonistas()
            atualizar_dropdowns_equipe()
            modal_excluir.open = False
            mostrar_aviso(f"'{nome_rem}' removido da lista.", "#DC2626")

    modal_excluir = ft.AlertDialog(
        title=ft.Text("➖ Excluir Plantonista"),
        content=dd_excluir_med,
        actions=[
            ft.TextButton("Cancelar", on_click=lambda e: setattr(modal_excluir, "open", False) or page.update()),
            ft.ElevatedButton("Excluir", bgcolor="#DC2626", color="white", on_click=confirmar_exclusao)
        ]
    )

    def abrir_modal_incluir(e):
        page.dialog = modal_incluir
        modal_incluir.open = True
        page.update()

    def abrir_modal_excluir(e):
        dd_excluir_med.options = [ft.dropdown.Option(p) for p in plantonistas]
        dd_excluir_med.value = plantonistas[0] if plantonistas else None
        page.dialog = modal_excluir
        modal_excluir.open = True
        page.update()

    def zerar_mes_click(e):
        nonlocal registros
        mes_nome = dd_mes.value
        ano = int(dd_ano.value)
        registros = [r for r in registros if not (r["mes"] == mes_nome and r["ano"] == ano)]
        estado["dia_atual"] = 1
        salvar_dados()
        atualizar_info_dia()
        atualizar_lista_historico()
        mostrar_aviso(f"Mês {mes_nome}/{ano} zerado!", "#475569")

    # ==========================================================================
    # GERAÇÃO DO EXCEL COMPLETO (ABA 1 DETALHADA + ABA 2 RESUMO, SEM COLUNA MÊS)
    # ==========================================================================
    def gerar_excel_click(e):
        mes_filtro = dd_mes.value
        ano_filtro = int(dd_ano.value)
        registros_mes = [r for r in registros if r["mes"] == mes_filtro and r["ano"] == ano_filtro]

        if not registros_mes:
            mostrar_aviso(f"Não há plantões salvos em {mes_filtro}/{ano_filtro}!", "#DC2626")
            return

        wb = openpyxl.Workbook()

        fonte_titulo = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        fonte_dados = Font(name="Calibri", size=11)
        fonte_subtotal = Font(name="Calibri", size=11, bold=True, color="1E3A8A")

        fill_titulo = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        fill_cabecalho = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        fill_dia = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        fill_noite = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        fill_subtotal = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")

        borda = Border(
            left=Side(style="thin", color="CBD5E1"),
            right=Side(style="thin", color="CBD5E1"),
            top=Side(style="thin", color="CBD5E1"),
            bottom=Side(style="thin", color="CBD5E1")
        )
        borda_subtotal = Border(
            left=Side(style="thin", color="93C5FD"),
            right=Side(style="thin", color="93C5FD"),
            top=Side(style="medium", color="2563EB"),
            bottom=Side(style="double", color="1E3A8A")
        )

        alin_centro = Alignment(horizontal="center", vertical="center")
        alin_esq = Alignment(horizontal="left", vertical="center")
        alin_dir = Alignment(horizontal="right", vertical="center")

        nomes_no_mes = sorted(
            set(plantonistas).union({r["nome"] for r in registros_mes}),
            key=lambda s: s.upper()
        )

        # ABA 1: DETALHADO POR PLANTONISTA (SEM COLUNA MÊS)
        ws_colab = wb.active
        ws_colab.title = "Detalhado por Plantonista"

        ws_colab.merge_cells("A1:F1")
        ws_colab["A1"] = f"RELATÓRIO DE PLANTÕES POR COLABORADOR - {mes_filtro} / {ano_filtro}"
        ws_colab["A1"].font = fonte_titulo
        ws_colab["A1"].fill = fill_titulo
        ws_colab["A1"].alignment = alin_centro
        ws_colab.row_dimensions[1].height = 28

        cabecalhos = ["Nome", "Dia", "Noite/Dia", "Modalidade", "Total Valor Dia", "Total Valor Noite"]
        ws_colab.row_dimensions[2].height = 25
        for col_idx, texto in enumerate(cabecalhos, start=1):
            cel = ws_colab.cell(row=2, column=col_idx, value=texto)
            cel.font = fonte_cabecalho
            cel.fill = fill_cabecalho
            cel.alignment = alin_centro
            cel.border = borda

        linha_atual = 3
        for nome_med in nomes_no_mes:
            regs_med = [r for r in registros_mes if r["nome"] == nome_med]
            if not regs_med:
                continue

            plantoes_dia = sorted([r for r in regs_med if r["noite_dia"].startswith("Dia")], key=lambda x: x["dia_num"])
            plantoes_noite = sorted([r for r in regs_med if r["noite_dia"].startswith("Noite")], key=lambda x: x["dia_num"])

            soma_dia_med = 0.0
            soma_noite_med = 0.0

            for reg in plantoes_dia:
                ws_colab.row_dimensions[linha_atual].height = 21
                soma_dia_med += reg["valor_dia"]
                valores = [reg["nome"], reg["dia_str"], reg["noite_dia"], reg["modalidade"], reg["valor_dia"], 0.0]
                for col_idx, val in enumerate(valores, start=1):
                    cel = ws_colab.cell(row=linha_atual, column=col_idx, value=val)
                    cel.font = fonte_dados
                    cel.fill = fill_dia
                    cel.border = borda
                    if col_idx == 1:
                        cel.alignment = alin_esq
                    elif col_idx in (5, 6):
                        cel.alignment = alin_dir
                        cel.number_format = 'R$ #,##0.00'
                    else:
                        cel.alignment = alin_centro
                linha_atual += 1

            for reg in plantoes_noite:
                ws_colab.row_dimensions[linha_atual].height = 21
                soma_noite_med += reg["valor_noite"]
                valores = [reg["nome"], reg["dia_str"], reg["noite_dia"], reg["modalidade"], 0.0, reg["valor_noite"]]
                for col_idx, val in enumerate(valores, start=1):
                    cel = ws_colab.cell(row=linha_atual, column=col_idx, value=val)
                    cel.font = fonte_dados
                    cel.fill = fill_noite
                    cel.border = borda
                    if col_idx == 1:
                        cel.alignment = alin_esq
                    elif col_idx in (5, 6):
                        cel.alignment = alin_dir
                        cel.number_format = 'R$ #,##0.00'
                    else:
                        cel.alignment = alin_centro
                linha_atual += 1

            ws_colab.row_dimensions[linha_atual].height = 25
            total_geral_med = soma_dia_med + soma_noite_med
            resumo_turnos = f"{len(plantoes_dia)} Dia | {len(plantoes_noite)} Noite"
            total_med_fmt = formatar_moeda(total_geral_med)

            linha_total_med = [
                f"TOTAL - {nome_med}",
                resumo_turnos,
                "Soma Dia / Noite",
                f"Total: {total_med_fmt}",
                soma_dia_med,
                soma_noite_med
            ]
            for col_idx, val in enumerate(linha_total_med, start=1):
                cel = ws_colab.cell(row=linha_atual, column=col_idx, value=val)
                cel.font = fonte_subtotal
                cel.fill = fill_subtotal
                cel.border = borda_subtotal
                if col_idx == 1:
                    cel.alignment = alin_esq
                elif col_idx in (5, 6):
                    cel.alignment = alin_dir
                    cel.number_format = 'R$ #,##0.00'
                else:
                    cel.alignment = alin_centro

            linha_atual += 2

        # ABA 2: QUADRO RESUMO POR PLANTONISTA (SEM COLUNA MÊS)
        ws_resumo = wb.create_sheet(title="Resumo por Plantonista")
        ws_resumo.merge_cells("A1:H1")
        ws_resumo["A1"] = f"QUADRO RESUMO POR PLANTONISTA - {mes_filtro} / {ano_filtro}"
        ws_resumo["A1"].font = fonte_titulo
        ws_resumo["A1"].fill = fill_titulo
        ws_resumo["A1"].alignment = alin_centro
        ws_resumo.row_dimensions[1].height = 28

        cabecalhos_resumo = [
            "Nome", "Qtd Plantões Dia", "Dias Trabalhados (Dia)", "Total Valor Dia",
            "Qtd Plantões Noite", "Dias Trabalhados (Noite)", "Total Valor Noite", "Total Geral Plantonista"
        ]
        ws_resumo.row_dimensions[2].height = 26
        for col_idx, texto in enumerate(cabecalhos_resumo, start=1):
            cel = ws_resumo.cell(row=2, column=col_idx, value=texto)
            cel.font = fonte_cabecalho
            cel.fill = fill_cabecalho
            cel.alignment = alin_centro
            cel.border = borda

        linha_res = 3
        for idx, nome_med in enumerate(nomes_no_mes):
            regs_med = [r for r in registros_mes if r["nome"] == nome_med]
            plantoes_dia = sorted([r for r in regs_med if r["noite_dia"].startswith("Dia")], key=lambda x: x["dia_num"])
            plantoes_noite = sorted([r for r in regs_med if r["noite_dia"].startswith("Noite")], key=lambda x: x["dia_num"])

            dias_str_dia = ", ".join([f"{r['dia_num']:02d}" for r in plantoes_dia]) if plantoes_dia else "-"
            dias_str_noite = ", ".join([f"{r['dia_num']:02d}" for r in plantoes_noite]) if plantoes_noite else "-"

            soma_dia = sum(r["valor_dia"] for r in plantoes_dia)
            soma_noite = sum(r["valor_noite"] for r in plantoes_noite)
            soma_total_med = soma_dia + soma_noite

            valores_res = [
                nome_med, len(plantoes_dia), dias_str_dia, soma_dia,
                len(plantoes_noite), dias_str_noite, soma_noite, soma_total_med
            ]
            ws_resumo.row_dimensions[linha_res].height = 22
            for col_idx, val in enumerate(valores_res, start=1):
                cel = ws_resumo.cell(row=linha_res, column=col_idx, value=val)
                cel.font = fonte_subtotal if col_idx == 8 else fonte_dados
                cel.border = borda
                if idx % 2 == 1:
                    cel.fill = fill_noite
                if col_idx == 1:
                    cel.alignment = alin_esq
                elif col_idx in (4, 7, 8):
                    cel.alignment = alin_dir
                    cel.number_format = 'R$ #,##0.00'
                else:
                    cel.alignment = alin_centro
            linha_res += 1

        for ws in [ws_colab, ws_resumo]:
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.row == 1:
                        continue
                    val_str = str(cell.value or "")
                    if len(val_str) > max_len:
                        max_len = len(val_str)
                ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

        nome_arq = f"Plantoes_Por_Colaborador_{mes_filtro}_{ano_filtro}.xlsx"
        caminho_temp = os.path.join(tempfile.gettempdir(), nome_arq)
        wb.save(caminho_temp)
        estado["caminho_temp_excel"] = caminho_temp

        # Tenta gravar também em Download do Android caso tenha permissão direta, e abre o seletor nativo
        try:
            pasta_dl = "/storage/emulated/0/Download"
            if os.path.exists(pasta_dl):
                wb.save(os.path.join(pasta_dl, nome_arq))
                mostrar_aviso(f"Salvo em Download/{nome_arq}!", "#16A34A")
                return
        except Exception:
            pass

        seletor_arquivo.save_file(
            dialog_title="Salvar Planilha de Plantões",
            file_name=nome_arq,
            allowed_extensions=["xlsx"]
        )

    # ==========================================================================
    # LAYOUT NA TELA
    # ==========================================================================
    page.add(
        ft.Container(
            bgcolor="#0F172A",
            padding=14,
            border_radius=10,
            content=ft.Column(
                spacing=8,
                controls=[
                    ft.Text("🏥 ESCALA HOSPITALAR", color="white", size=16, weight=ft.FontWeight.BOLD),
                    ft.Row(
                        controls=[
                            ft.ElevatedButton("➕ Incluir", bgcolor="#10B981", color="white", on_click=abrir_modal_incluir, expand=True),
                            ft.ElevatedButton("➖ Excluir", bgcolor="#EF4444", color="white", on_click=abrir_modal_excluir, expand=True),
                        ]
                    )
                ]
            )
        ),
        ft.Container(
            bgcolor="white",
            padding=12,
            border_radius=10,
            border=ft.border.all(1, "#CBD5E1"),
            content=ft.Column(
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=8,
                controls=[
                    ft.Row(controls=[dd_mes, dd_ano]),
                    ft.Row(
                        alignment=ft.MainAxisAlignment.CENTER,
                        controls=[
                            ft.IconButton(icon=ft.icons.CHEVRON_LEFT, on_click=voltar_dia, bgcolor="#E2E8F0"),
                            ft.Container(
                                bgcolor="#DBEAFE",
                                padding=ft.padding.symmetric(horizontal=18, vertical=6),
                                border_radius=8,
                                content=txt_dia_atual
                            ),
                            ft.IconButton(icon=ft.icons.CHEVRON_RIGHT, on_click=avancar_dia, bgcolor="#E2E8F0"),
                        ]
                    ),
                    txt_dia_semana
                ]
            )
        ),
        ft.Container(
            bgcolor="white",
            padding=12,
            border_radius=10,
            border=ft.border.all(1, "#CBD5E1"),
            content=ft.Column(
                spacing=6,
                controls=[
                    ft.Container(
                        bgcolor="#FEF3C7", padding=6, border_radius=6,
                        content=ft.Text("☀️ MANHÃ (07h - 19h)", weight=ft.FontWeight.BOLD, color="#92400E")
                    ),
                    dd_manha,
                    dd_manha_2,
                    ft.Row(wrap=True, controls=list(chks_manha.values()))
                ]
            )
        ),
        ft.Container(
            bgcolor="white",
            padding=12,
            border_radius=10,
            border=ft.border.all(1, "#CBD5E1"),
            content=ft.Column(
                spacing=6,
                controls=[
                    ft.Container(
                        bgcolor="#E0E7FF", padding=6, border_radius=6,
                        content=ft.Text("🌙 NOITE (19h - 07h)", weight=ft.FontWeight.BOLD, color="#1E40AF")
                    ),
                    dd_noite,
                    dd_noite_2,
                    ft.Row(wrap=True, controls=list(chks_noite.values()))
                ]
            )
        ),
        ft.Row(
            controls=[
                ft.ElevatedButton("💾 Salvar Dia", bgcolor="#2563EB", color="white", height=48, expand=True, on_click=salvar_dia_click),
                ft.ElevatedButton("📊 Gerar Planilha", bgcolor="#16A34A", color="white", height=48, expand=True, on_click=gerar_excel_click),
            ]
        ),
        ft.Container(
            bgcolor="white",
            padding=12,
            border_radius=10,
            border=ft.border.all(1, "#CBD5E1"),
            content=ft.Column(
                controls=[
                    ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        controls=[txt_resumo_mes, ft.TextButton("⚠️ Zerar Mês", on_click=zerar_mes_click)]
                    ),
                    coluna_historico
                ]
            )
        )
    )

    ao_mudar_mes_ano()


if __name__ == "__main__":
    ft.app(target=main)