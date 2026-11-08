# -*- coding: utf-8 -*-
"""
Balança Lartec · leitura direta por Bluetooth, sem o app OKOK
Script para o Pythonista 3 no iPhone.        Versão 1.1  (07/11/2026)

COMO USAR
  1. Rode este script (botão ▶). Na primeira vez, permita o Bluetooth para o Pythonista.
  2. Suba na balança. O peso aparece ao vivo e, quando estabiliza, é salvo sozinho.
  3. O histórico fica no arquivo  pesos.csv,  na mesma pasta deste script
     (abre no Excel/Numbers e pode ser importado na página "Balança Lartec").
  4. Opcional: ligue "Enviar ao app Saúde". Cada pesagem salva chama o atalho
     "Registrar Peso" (app Atalhos), que grava o peso no Saúde.

COMO FUNCIONA
  A balança (compatível com o app OKOK International) não precisa de conexão: ela
  transmite o peso dentro dos anúncios Bluetooth ("manufacturer data"). O script
  escuta esses anúncios pelo CoreBluetooth do iOS (via objc_util) e decodifica o
  formato OKOK, com base no OkOkHandler do projeto openScale (GPL-3.0), conferido
  com a própria balança Lartec: ID do fabricante terminado em C0, peso nos 2
  primeiros bytes, byte 6 com estabilidade/unidade/casas decimais.
"""

import os
import re
import io
import json
import time
import datetime
import collections

try:                                    # módulos do Pythonista (iOS)
    import ui
    import console
except ImportError:                     # fora do iPhone (testes)
    ui = None
    console = None

VERSAO = '1.1'
PASTA = os.path.dirname(os.path.abspath(__file__))
ARQUIVO = os.path.join(PASTA, 'pesos.csv')
CONFIG = os.path.join(PASTA, 'balanca_config.json')
ATALHO_SAUDE = 'Registrar Peso'      # nome do atalho no app Atalhos

# cores (tema escuro)
BG, CARD, LINHA = '#0f1115', '#181b21', '#2a2f37'
TINTA, APAGADO = '#e8eaed', '#9aa1ab'
DESTAQUE, VERDE, VERMELHO, AMARELO = '#7c9cff', '#4ade80', '#f87171', '#fbbf24'

DIAS = ['seg', 'ter', 'qua', 'qui', 'sex', 'sáb', 'dom']


# ════════════════════════════════════════════════════════════════════
#  1. Decodificação do anúncio da balança (formatos OKOK)
# ════════════════════════════════════════════════════════════════════

def _converter_unidade(bruto, msb, lsb, props):
    """Aplica casas decimais e unidade indicadas no byte de propriedades."""
    divisor = (10.0, 1.0, 100.0, 10.0)[(props >> 1) & 3]
    unidade = (props >> 3) & 3
    if unidade == 0:                                  # kg
        return bruto / divisor
    if unidade == 1:                                  # jin (meio quilo)
        return bruto / divisor / 2
    if unidade == 2:                                  # libras
        return bruto / divisor / 2.204623
    return msb * 6.350293 + (lsb / divisor) * 0.453592  # stone:libra


def decodificar(dados):
    """Recebe o "manufacturer data" completo (2 bytes do ID do fabricante,
    little-endian, + carga útil) e devolve um dicionário com a leitura, ou None.

    Chaves: kg, estavel, tem_flag (se o formato informa estabilidade),
            imp (impedância em ohm), formato, id_fab, id (identificação da balança)
    """
    if not dados or len(dados) < 4:
        return None
    dados = bytes(dados)
    id_fab = dados[0] | (dados[1] << 8)
    d = dados[2:]

    def u16(i, j):
        return (d[i] << 8) | d[j]

    leitura = None
    if id_fab == 0x20CA and len(d) == 19:                       # OKOK V20
        soma = 0x20
        for b in d[:12]:
            soma ^= b
        if soma & 0xFF == d[12]:
            flags = d[6]
            divisor = 100.0 if flags & 0x04 else 10.0
            leitura = dict(kg=u16(8, 9) / divisor, estavel=bool(flags & 1),
                           tem_flag=True, imp=u16(10, 11) / 10.0, formato='V20', id=None)
    elif id_fab == 0x11CA and len(d) == 23:                     # OKOK V11
        soma = 0xCA ^ 0x11
        for b in d[:16]:
            soma ^= b
        if soma & 0xFF == d[16]:
            leitura = dict(kg=_converter_unidade(u16(3, 4), d[3], d[4], d[9]), estavel=True,
                           tem_flag=False, imp=0.0, formato='V11', id=None)
    elif id_fab == 0xF0FF and len(d) >= 4:                      # OKOK VF0
        leitura = dict(kg=u16(3, 2) / 10.0, estavel=True, tem_flag=False,
                       imp=0.0, formato='VF0', id=None)
    elif (id_fab & 0xFF) == 0xC0 and len(d) >= 13:              # OKOK C0 (a Lartec)
        attr = d[6]
        leitura = dict(kg=_converter_unidade(u16(0, 1), d[0], d[1], attr),
                       estavel=bool(attr & 1), tem_flag=True, imp=u16(2, 3) / 10.0,
                       formato='C0', id=d[7:13].hex())
    if leitura is None or not (0.0 <= leitura['kg'] < 300.0):
        return None
    leitura['kg'] = round(leitura['kg'], 2)
    leitura['id_fab'] = id_fab
    return leitura


# ════════════════════════════════════════════════════════════════════
#  2. Quando gravar
# ════════════════════════════════════════════════════════════════════

class Sessao(object):
    """Decide o que fazer com cada leitura.

    Regra principal: um peso estável só é gravado sozinho se, antes dele, a
    balança mostrou a medição em andamento (peso instável). Ao acordar, a balança
    volta a anunciar o último peso medido; essa regra impede gravar esse peso
    antigo de novo. Se a medição em andamento não foi captada, a tela oferece
    o botão "Salvar" para gravar à mão.

    Retornos de processar(): 'vazia', 'medindo', 'salvar', 'repetido', 'antigo'
    """

    def __init__(self, ultimo_salvo=None):
        self.armado = False
        self.ultimo_salvo = ultimo_salvo
        self.salvo_em = 0.0
        self.pendente = None
        self._recentes = []          # (t, kg) — para formatos sem indicação de estabilidade

    def _estavel_sem_flag(self, kg, agora):
        self._recentes.append((agora, kg))
        self._recentes = [(t, k) for t, k in self._recentes if agora - t <= 3.0]
        return (len(self._recentes) >= 3
                and all(abs(k - kg) < 0.01 for _, k in self._recentes)
                and agora - self._recentes[0][0] >= 1.5)

    def processar(self, leitura, agora):
        kg = leitura['kg']
        estavel = leitura['estavel'] if leitura['tem_flag'] else self._estavel_sem_flag(kg, agora)
        if kg < 2.0:
            return 'vazia'
        if not estavel:
            self.armado = True
            self.pendente = None
            return 'medindo'
        if self.armado:
            self.armado = False
            self.pendente = None
            # mesma pesagem medida de novo logo em seguida: não duplica
            if (self.ultimo_salvo is not None and agora - self.salvo_em < 90
                    and abs(kg - self.ultimo_salvo) < 0.5):
                return 'repetido'
            self.ultimo_salvo = kg
            self.salvo_em = agora
            return 'salvar'
        if self.ultimo_salvo is not None and abs(kg - self.ultimo_salvo) < 0.05:
            self.pendente = None
            return 'repetido'
        self.pendente = leitura
        return 'antigo'

    def confirmar_manual(self, kg, agora=None):
        self.pendente = None
        self.ultimo_salvo = kg
        self.salvo_em = agora or time.time()


# ════════════════════════════════════════════════════════════════════
#  3. Histórico em CSV
# ════════════════════════════════════════════════════════════════════

def _ler_data(texto):
    texto = texto.strip().strip('"')
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d',
                '%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M', '%d/%m/%Y'):
        try:
            return datetime.datetime.strptime(texto, fmt)
        except ValueError:
            pass
    return None


def _ler_numero(texto):
    texto = texto.strip().strip('"').lower().replace('kg', '').replace('ohm', '').strip()
    if re.match(r'^-?\d+,\d+$', texto):
        texto = texto.replace(',', '.')
    try:
        return float(texto)
    except ValueError:
        return None


def fmt_kg(v, casas=2):
    return ('%.*f' % (casas, v)).replace('.', ',')


def fmt_delta(v):
    if v is None:
        return '—'
    sinal = '+' if v > 0.004 else ('−' if v < -0.004 else '±')
    return sinal + fmt_kg(abs(v))


class Historico(object):
    """Lista de pesagens [(datetime, kg, impedancia)], gravada em CSV
    (separador ; e vírgula decimal, como o Excel brasileiro espera)."""

    CABECALHO = 'data;peso_kg;impedancia_ohm'

    def __init__(self, caminho):
        self.caminho = caminho
        self.itens = []
        self.carregar()

    def carregar(self):
        self.itens = []
        if not os.path.exists(self.caminho):
            return
        with open(self.caminho, encoding='utf-8-sig') as f:
            for linha in f:
                partes = re.split(r'[;\t]', linha.strip())
                if len(partes) < 2:
                    continue
                quando, kg = _ler_data(partes[0]), _ler_numero(partes[1])
                if quando is None or kg is None or not (10 < kg < 400):
                    continue
                imp = _ler_numero(partes[2]) if len(partes) > 2 else None
                self.itens.append((quando, round(kg, 2), imp or 0.0))
        self.itens.sort(key=lambda x: x[0])

    def gravar(self):
        tmp = self.caminho + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            f.write(self.CABECALHO + '\n')
            for quando, kg, imp in self.itens:
                f.write('%s;%s;%s\n' % (quando.strftime('%Y-%m-%d %H:%M'),
                                        fmt_kg(kg), fmt_kg(imp, 1)))
        os.replace(tmp, self.caminho)

    def adicionar(self, kg, imp=0.0, quando=None):
        quando = (quando or datetime.datetime.now()).replace(microsecond=0)
        self.itens.append((quando, round(kg, 2), imp or 0.0))
        self.itens.sort(key=lambda x: x[0])
        self.gravar()

    def apagar(self, indice):
        del self.itens[indice]
        self.gravar()

    def ultimo(self):
        return self.itens[-1] if self.itens else None

    def variacao(self, dias):
        """Diferença entre o último peso e o peso de 'dias' atrás (ou o mais próximo antes)."""
        if len(self.itens) < 2:
            return None
        ultimo = self.itens[-1]
        limite = ultimo[0] - datetime.timedelta(days=dias)
        referencia = None
        for item in self.itens:
            if item[0] <= limite:
                referencia = item
            else:
                break
        return None if referencia is None else ultimo[1] - referencia[1]


def grafico_png(itens, largura_px=1000):
    """Gera um PNG (bytes) com a evolução do peso e a média de 7 dias."""
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    import matplotlib.dates as mdates
    from matplotlib.ticker import FuncFormatter

    fig = Figure(figsize=(largura_px / 160.0, largura_px / 160.0 * 0.62), dpi=160, facecolor=BG)
    FigureCanvasAgg(fig)
    ax = fig.add_subplot(111, facecolor=BG)
    datas = [i[0] for i in itens]
    pesos = [i[1] for i in itens]
    media = []
    for i, (d, _) in enumerate(zip(datas, pesos)):
        janela = [p for dd, p in zip(datas[:i + 1], pesos[:i + 1])
                  if (d - dd).total_seconds() < 7 * 86400]
        media.append(sum(janela) / len(janela))
    ax.plot(datas, pesos, color=DESTAQUE, linewidth=1.8, marker='o', markersize=3.5, label='Peso')
    if len(itens) >= 3:
        ax.plot(datas, media, color=AMARELO, linewidth=1.6, linestyle='--', label='Média 7 dias')
    for lado in ('top', 'right'):
        ax.spines[lado].set_visible(False)
    for lado in ('left', 'bottom'):
        ax.spines[lado].set_color(LINHA)
    ax.tick_params(colors=APAGADO, labelsize=8)
    ax.grid(True, color=LINHA, linewidth=0.6)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m'))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: fmt_kg(v, 1)))
    ax.set_ylabel('kg', color=APAGADO, fontsize=9)
    legenda = ax.legend(loc='best', fontsize=8, frameon=False)
    for t in legenda.get_texts():
        t.set_color(APAGADO)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format='png', facecolor=BG)
    return buf.getvalue()


def ler_config():
    try:
        with open(CONFIG, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def gravar_config(cfg):
    with open(CONFIG, 'w', encoding='utf-8') as f:
        json.dump(cfg, f)


def url_saude(kg):
    """URL que roda o atalho do Saúde passando o peso em GRAMAS (número inteiro,
    para não depender de vírgula ou ponto decimal) e volta ao Pythonista."""
    from urllib.parse import quote
    return ('shortcuts://x-callback-url/run-shortcut?name=%s&input=text&text=%d&x-success=%s'
            % (quote(ATALHO_SAUDE, safe=''), int(round(kg * 1000)), quote('pythonista3://', safe='')))


def enviar_saude(kg):
    try:
        import shortcuts
        shortcuts.open_url(url_saude(kg))
    except ImportError:
        import webbrowser
        webbrowser.open(url_saude(kg))


# ════════════════════════════════════════════════════════════════════
#  4. Bluetooth (CoreBluetooth via objc_util)
# ════════════════════════════════════════════════════════════════════

_ble = {}                       # mantém vivas as referências Objective-C
ESTADOS_BT = {0: 'desconhecido', 1: 'reiniciando', 2: 'sem_suporte',
              3: 'sem_permissao', 4: 'desligado', 5: 'ligado'}


def ble_iniciar(app):
    from objc_util import ObjCClass, ObjCInstance, create_objc_class, nsdata_to_bytes, ns

    def centralManagerDidUpdateState_(_self, _cmd, central):
        try:
            c = ObjCInstance(central)
            estado = ESTADOS_BT.get(int(c.state()), 'desconhecido')
            if estado == 'ligado':
                # AllowDuplicates: recebe TODOS os anúncios (o peso muda a cada um)
                c.scanForPeripheralsWithServices_options_(
                    None, ns({'kCBScanOptionAllowDuplicates': True}))
            app.ao_mudar_bluetooth(estado)
        except Exception as e:
            app.registrar('erro (estado): %r' % (e,))

    def centralManager_didDiscoverPeripheral_advertisementData_RSSI_(_self, _cmd, central,
                                                                     periferico, anuncio, rssi):
        try:
            dic = ObjCInstance(anuncio)
            fab = dic.objectForKey_('kCBAdvDataManufacturerData')
            if fab is None:
                return
            sinal = int(ObjCInstance(rssi).intValue()) if rssi else 0
            app.ao_receber(nsdata_to_bytes(fab), sinal)
        except Exception as e:
            app.registrar('erro (anúncio): %r' % (e,))

    Delegado = create_objc_class(
        'BalancaLartecBLE',
        methods=[centralManagerDidUpdateState_,
                 centralManager_didDiscoverPeripheral_advertisementData_RSSI_],
        protocols=['CBCentralManagerDelegate'])
    delegado = Delegado.alloc().init()
    try:
        CBCentralManager = ObjCClass('CBCentralManager')
    except ValueError:
        import ctypes
        ctypes.cdll.LoadLibrary('/System/Library/Frameworks/CoreBluetooth.framework/CoreBluetooth')
        CBCentralManager = ObjCClass('CBCentralManager')
    central = CBCentralManager.alloc().initWithDelegate_queue_(delegado, None)
    _ble.update(delegado=delegado, central=central, classe=Delegado)


def ble_parar():
    central = _ble.get('central')
    if central is not None:
        try:
            central.stopScan()
            central.setDelegate_(None)
        except Exception:
            pass
    _ble.clear()


# ════════════════════════════════════════════════════════════════════
#  5. Tela
# ════════════════════════════════════════════════════════════════════

def _rotulo(texto='', tamanho=15, cor=TINTA, negrito=False, alinhamento=None, linhas=1):
    r = ui.Label()
    r.text = texto
    r.number_of_lines = linhas
    r.text_color = cor
    r.font = ('<System-Bold>' if negrito else '<System>', tamanho)
    r.alignment = ui.ALIGN_CENTER if alinhamento is None else alinhamento
    return r


def _botao(titulo, acao, cor_fundo=CARD, cor_texto=TINTA, tamanho=15):
    b = ui.Button()
    b.title = titulo
    b.action = acao
    b.background_color = cor_fundo
    b.tint_color = cor_texto
    b.font = ('<System-Bold>', tamanho)
    b.corner_radius = 10
    return b


class FonteTabela(object):
    """Fonte de dados da lista de pesagens (mais recente no topo)."""

    def __init__(self, app):
        self.app = app

    def _item(self, linha):
        itens = self.app.hist.itens
        indice = len(itens) - 1 - linha
        return indice, itens[indice], (itens[indice - 1] if indice > 0 else None)

    def tableview_number_of_sections(self, tv):
        return 1

    def tableview_number_of_rows(self, tv, secao):
        return len(self.app.hist.itens)

    def tableview_title_for_header(self, tv, secao):
        n = len(self.app.hist.itens)
        return 'Pesagens (%d) — deslize para apagar' % n if n else 'Nenhuma pesagem ainda'

    def tableview_cell_for_row(self, tv, secao, linha):
        _, (quando, kg, _imp), anterior = self._item(linha)
        c = ui.TableViewCell('value1')
        c.background_color = BG
        c.selectable = False
        c.text_label.text = '%s %s' % (DIAS[quando.weekday()], quando.strftime('%d/%m/%Y  %H:%M'))
        c.text_label.text_color = TINTA
        c.text_label.font = ('<System>', 15)
        delta = '' if anterior is None else '   ' + fmt_delta(kg - anterior[1])
        c.detail_text_label.text = '%s kg%s' % (fmt_kg(kg), delta)
        c.detail_text_label.text_color = APAGADO if anterior is None else (
            VERMELHO if kg - anterior[1] > 0.05 else VERDE if kg - anterior[1] < -0.05 else APAGADO)
        c.detail_text_label.font = ('Menlo', 15)
        return c

    def tableview_can_delete(self, tv, secao, linha):
        return True

    def tableview_delete(self, tv, secao, linha):
        indice, _, _ = self._item(linha)
        self.app.hist.apagar(indice)
        ultimo = self.app.hist.ultimo()
        self.app.sessao.ultimo_salvo = ultimo[1] if ultimo else None
        self.app.atualizar_lista()


class App(ui.View if ui else object):

    def __init__(self):
        self.name = 'Balança Lartec · v' + VERSAO
        self.background_color = BG
        self.update_interval = 1.0
        self.hist = Historico(ARQUIVO)
        ultimo = self.hist.ultimo()
        self.sessao = Sessao(ultimo[1] if ultimo else None)
        self.estado_bt = 'iniciando'
        self.visto_em = 0.0
        self.n_anuncios = 0
        self.ultimo_bruto = None
        self.log = collections.deque(maxlen=80)
        self.tela_diag = None
        self.leitura_atual = None
        self.fab_vistos = set()
        self.cfg = ler_config()

        self.lbl_status = _rotulo('Iniciando o Bluetooth…', 14, APAGADO, linhas=2)
        self.lbl_peso = _rotulo('--,-- kg', 62, APAGADO, True)
        self.lbl_info = _rotulo('Suba na balança', 16, APAGADO, linhas=2)
        self.btn_salvar = _botao('Salvar este peso', self.salvar_manual, DESTAQUE, 'white', 17)
        self.btn_salvar.hidden = True
        self.lbl_resumo = _rotulo('', 14, APAGADO)
        self.tabela = ui.TableView()
        self.tabela.background_color = BG
        self.tabela.separator_color = LINHA
        self.tabela.row_height = 46
        self.fonte = FonteTabela(self)
        self.tabela.data_source = self.fonte
        self.tabela.delegate = self.fonte
        self.lbl_saude = _rotulo('Enviar ao app Saúde', 15, TINTA, alinhamento=0)
        self.sw_saude = ui.Switch()
        self.sw_saude.value = bool(self.cfg.get('saude'))
        self.sw_saude.action = self.mudar_saude
        self.btn_grafico = _botao('Gráfico', self.mostrar_grafico)
        self.btn_exportar = _botao('Exportar CSV', self.exportar)
        self.btn_diag = _botao('Diagnóstico', self.mostrar_diagnostico)
        for v in (self.lbl_status, self.lbl_peso, self.lbl_info, self.btn_salvar,
                  self.lbl_resumo, self.tabela, self.lbl_saude, self.sw_saude,
                  self.btn_grafico, self.btn_exportar, self.btn_diag):
            self.add_subview(v)
        self.atualizar_lista()

    # ── layout ──────────────────────────────────────────────────────
    def layout(self):
        w, h, m = self.width, self.height, 16
        y = 8
        self.lbl_status.frame = (m, y, w - 2 * m, 36)
        y += 38
        self.lbl_peso.frame = (m, y, w - 2 * m, 80)
        y += 84
        self.lbl_info.frame = (m, y, w - 2 * m, 40)
        y += 44
        self.btn_salvar.frame = (m + 30, y, w - 2 * m - 60, 44)
        y += 54
        self.lbl_resumo.frame = (m, y, w - 2 * m, 22)
        y += 28
        base = h - 44 - 40          # deixa espaço para a barra inferior do iPhone
        self.tabela.frame = (0, y, w, base - y - 52)
        self.lbl_saude.frame = (m, base - 42, w - 2 * m - 60, 32)
        self.sw_saude.frame = (w - m - 51, base - 41, 51, 31)
        bw = (w - 2 * m - 16) / 3.0
        for i, b in enumerate((self.btn_grafico, self.btn_exportar, self.btn_diag)):
            b.frame = (m + i * (bw + 8), base, bw, 44)

    # ── eventos do Bluetooth ────────────────────────────────────────
    def ao_mudar_bluetooth(self, estado):
        self.estado_bt = estado
        self.registrar('Bluetooth: ' + estado)
        self.atualizar_status()

    def ao_receber(self, dados, rssi):
        self.n_anuncios += 1
        leitura = decodificar(dados)
        if leitura is None:
            dados = bytes(dados)
            if len(dados) >= 2 and rssi >= -65:
                fab = dados[0] | (dados[1] << 8)
                if fab not in self.fab_vistos:
                    self.fab_vistos.add(fab)
                    self.registrar('outro aparelho fab=0x%04X rssi %d: %s' % (fab, rssi, dados.hex()))
            return
        agora = time.time()
        self.visto_em = agora
        self.leitura_atual = leitura
        bruto = bytes(dados).hex()
        acao = self.sessao.processar(leitura, agora)
        if bruto != self.ultimo_bruto:            # registra só quando o anúncio muda
            self.ultimo_bruto = bruto
            self.registrar('%s %s rssi %d → %s kg %s [%s]' % (
                leitura['formato'], ' '.join(bruto[i:i + 4] for i in range(0, len(bruto), 4)),
                rssi, fmt_kg(leitura['kg']), 'estável' if leitura['estavel'] else 'medindo', acao))
        self.mostrar_leitura(leitura, acao)

    def mostrar_leitura(self, leitura, acao):
        kg = leitura['kg']
        self.btn_salvar.hidden = acao != 'antigo'
        if acao == 'vazia':
            self.lbl_peso.text, self.lbl_peso.text_color = '--,-- kg', APAGADO
            self.lbl_info.text, self.lbl_info.text_color = 'Balança acordada — suba nela', APAGADO
        elif acao == 'medindo':
            self.lbl_peso.text, self.lbl_peso.text_color = fmt_kg(kg) + ' kg', TINTA
            self.lbl_info.text, self.lbl_info.text_color = 'Medindo… fique parado', AMARELO
        elif acao == 'salvar':
            self.hist.adicionar(kg, leitura['imp'])
            self.lbl_peso.text, self.lbl_peso.text_color = fmt_kg(kg) + ' kg', VERDE
            self.lbl_info.text = '✓ Salvo às ' + datetime.datetime.now().strftime('%H:%M')
            self.lbl_info.text_color = VERDE
            self.registrar('SALVO %s kg' % fmt_kg(kg))
            self.atualizar_lista()
            self._som()
            self._saude(kg)
        elif acao == 'repetido':
            self.lbl_peso.text, self.lbl_peso.text_color = fmt_kg(kg) + ' kg', VERDE
            if not self.lbl_info.text.startswith('✓'):
                self.lbl_info.text, self.lbl_info.text_color = '✓ Já salvo', VERDE
        elif acao == 'antigo':
            self.lbl_peso.text, self.lbl_peso.text_color = fmt_kg(kg) + ' kg', APAGADO
            self.lbl_info.text = 'Peso anterior guardado na balança.\nSuba nela para pesar de novo.'
            self.lbl_info.text_color = APAGADO
            self.btn_salvar.title = 'Salvar %s kg' % fmt_kg(kg)
        self.atualizar_status()

    def salvar_manual(self, sender):
        leitura = self.sessao.pendente
        if not leitura:
            return
        self.hist.adicionar(leitura['kg'], leitura['imp'])
        self.sessao.confirmar_manual(leitura['kg'])
        self.btn_salvar.hidden = True
        self.lbl_peso.text_color = VERDE
        self.lbl_info.text, self.lbl_info.text_color = '✓ Salvo às ' + time.strftime('%H:%M'), VERDE
        self.lbl_peso.text = fmt_kg(leitura['kg']) + ' kg'
        self.registrar('SALVO (manual) %s kg' % fmt_kg(leitura['kg']))
        self.atualizar_lista()
        self._saude(leitura['kg'])

    # ── atualizações periódicas ─────────────────────────────────────
    def update(self):
        self.atualizar_status()

    def atualizar_status(self):
        e = self.estado_bt
        if e == 'sem_permissao':
            txt, cor = 'Sem permissão: Ajustes › Pythonista › ligue o Bluetooth', VERMELHO
        elif e == 'desligado':
            txt, cor = 'Bluetooth desligado — ligue na Central de Controle', VERMELHO
        elif e == 'sem_suporte':
            txt, cor = 'Este aparelho não tem Bluetooth LE', VERMELHO
        elif e != 'ligado':
            txt, cor = 'Iniciando o Bluetooth…', APAGADO
        elif not self.visto_em:
            txt, cor = '● Procurando a balança — pise nela para acordar', DESTAQUE
        else:
            seg = int(time.time() - self.visto_em)
            if seg <= 3:
                txt, cor = '● Balança conectada', VERDE
            else:
                txt, cor = '● Balança sem sinal há %d s (desligou?)' % seg, APAGADO
        self.lbl_status.text, self.lbl_status.text_color = txt, cor

    def atualizar_lista(self):
        self.tabela.reload_data()
        partes = []
        for dias in (7, 30):
            v = self.hist.variacao(dias)
            if v is not None:
                partes.append('%d dias: %s kg' % (dias, fmt_delta(v)))
        self.lbl_resumo.text = '   ·   '.join(partes)

    # ── app Saúde ───────────────────────────────────────────────────
    def mudar_saude(self, sender):
        self.cfg['saude'] = bool(sender.value)
        gravar_config(self.cfg)
        self.registrar('Enviar ao Saúde: %s' % ('sim' if sender.value else 'não'))

    def _saude(self, kg):
        if self.cfg.get('saude'):
            self.registrar('Saúde: chamando atalho "%s" com %s kg' % (ATALHO_SAUDE, fmt_kg(kg)))
            ui.delay(lambda: enviar_saude(kg), 1.2)     # deixa ver o "✓ Salvo" antes

    # ── botões ──────────────────────────────────────────────────────
    def _som(self):
        try:
            import sound
            sound.play_effect('arcade:Coin_1')
        except Exception:
            pass

    @staticmethod
    def _tela(titulo, conteudo):
        v = ui.View()
        v.name = titulo
        v.background_color = BG
        conteudo.frame = (0, 0, 400, 400)
        conteudo.flex = 'WH'
        v.add_subview(conteudo)
        v.present('sheet', title_bar_color=BG, title_color=TINTA)
        return v

    @ui.in_background if ui else (lambda f: f)
    def mostrar_grafico(self, sender):
        if len(self.hist.itens) < 2:
            console.hud_alert('O gráfico aparece a partir de 2 pesagens', 'error', 2)
            return
        console.show_activity()
        try:
            png = grafico_png(self.hist.itens)
        finally:
            console.hide_activity()
        img = ui.ImageView()
        img.image = ui.Image.from_data(png)
        img.content_mode = ui.CONTENT_SCALE_ASPECT_FIT
        img.background_color = BG
        self._tela('Evolução do peso', img)

    @ui.in_background if ui else (lambda f: f)
    def exportar(self, sender):
        if not os.path.exists(ARQUIVO):
            console.hud_alert('Ainda não há pesagens', 'error', 2)
            return
        console.open_in(ARQUIVO)

    def mostrar_diagnostico(self, sender):
        tv = ui.TextView()
        tv.editable = False
        tv.background_color = BG
        tv.text_color = TINTA
        tv.font = ('Menlo', 11)
        self.tela_diag = tv
        self._preencher_diag()
        self._tela('Diagnóstico', tv)

    def _preencher_diag(self):
        if self.tela_diag is None:
            return
        cab = ('Versão %s · Bluetooth: %s · anúncios recebidos: %d\n'
               'Arquivo: %s\n\n' % (VERSAO, self.estado_bt, self.n_anuncios, ARQUIVO))
        self.tela_diag.text = cab + '\n'.join(reversed(self.log))

    def registrar(self, texto):
        self.log.append(time.strftime('%H:%M:%S ') + texto)
        if self.tela_diag is not None and self.tela_diag.on_screen:
            self._preencher_diag()

    def will_close(self):
        ble_parar()


def main():
    app = App()
    app.present('fullscreen', title_bar_color=BG, title_color=TINTA, orientations=['portrait'])
    try:
        ble_iniciar(app)
    except Exception as e:
        app.registrar('Falha ao iniciar o Bluetooth: %r' % (e,))
        app.lbl_status.text = 'Falha ao iniciar o Bluetooth — veja Diagnóstico'
        app.lbl_status.text_color = VERMELHO


if __name__ == '__main__':
    main()
