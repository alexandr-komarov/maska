#!/usr/bin/env python3
"""Сборка синтетического набора «связанные документы одной сделки».

Запуск из корня репозитория:  python3 tests/deal/make.py
Пишет в tests/deal/ шесть-восемь файлов (docx и xlsx) и разметку tests/deal.json.
Сторонних пакетов нет: docx собираются из XML через zipfile, xlsx через
soffice --headless --convert-to xlsx из flat ODS (.fods во временной папке).
Все данные выдуманные. ИНН, ОГРН, ОГРНИП и ключ расчётного счёта считаются по настоящим алгоритмам.
"""
import json, os, shutil, subprocess, tempfile, zipfile
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))

# ---------------------------------------------------------------- реквизиты
def inn10(b9):
    k = [2, 4, 10, 3, 5, 9, 4, 6, 8]
    d = [int(c) for c in b9]
    return b9 + str(sum(a * b for a, b in zip(d, k)) % 11 % 10)

def inn12(b10):
    k1 = [7, 2, 4, 10, 3, 5, 9, 4, 6, 8]
    k2 = [3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8]
    d = [int(c) for c in b10]
    s = b10 + str(sum(a * b for a, b in zip(d, k1)) % 11 % 10)
    d = [int(c) for c in s]
    return s + str(sum(a * b for a, b in zip(d, k2)) % 11 % 10)

def ogrn(b12):
    return b12 + str(int(b12) % 11 % 10)

def ogrnip(b14):
    return b14 + str(int(b14) % 13 % 10)

def rs(prefix8, bik, tail11):
    """Расчётный счёт с верным ключом (веса 7,1,3 по последним трём цифрам БИК и 20 цифрам счёта)."""
    acc = prefix8 + '0' + tail11
    s = bik[-3:] + acc
    w = [7, 1, 3] * 8
    key = sum((int(c) * w[i]) % 10 for i, c in enumerate(s)) * 3 % 10
    return prefix8 + str(key) + tail11

BANKS = {
    'sber': ('ПАО Сбербанк', '044030653', '30101810500000000653', 'г. Санкт-Петербург'),
    'vtb': ('Банк ВТБ (ПАО)', '044030707', '30101810200000000707', 'г. Санкт-Петербург'),
    'alfa': ('АО «Альфа-Банк»', '044525593', '30101810200000000593', 'г. Москва'),
    'tink': ('АО «Тинькофф Банк»', '044525974', '30101810145250000974', 'г. Москва'),
    'kaz': ('Филиал «Казанский» банка', '049205805', '30101810600000000805', 'г. Казань'),
}

FIRMS = {
    'sv': dict(core='Северный Ветер Строй', short='ООО «Северный Ветер Строй»',
               full='Общество с ограниченной ответственностью «Северный Ветер Строй»',
               inn=inn10('781012345'), kpp='781001001', ogrn=ogrn('112784700123'),
               bank='sber', rs=rs('40702810', '044030653', '20000045671'),
               addr='ул. Балканская, д. 20, лит. А, офис 412', city='г. Санкт-Петербург',
               tel='+7 (812) 346-18-02', mail='office@severvetrostroy.example',
               dir='Нестерова Алла Викторовна', dir_gen='Нестеровой Аллы Викторовны',
               dir_ini='Нестерова А.В.', dir_pos='генерального директора'),
    'km': dict(core='Каменный Мост', short='ООО «Каменный Мост»',
               full='Общество с ограниченной ответственностью «Каменный Мост»',
               inn=inn10('781613579'), kpp='781601001', ogrn=ogrn('117784700456'),
               bank='vtb', rs=rs('40702810', '044030707', '30000087432'),
               addr='Лиговский пр., д. 87, корп. 2, пом. 14Н', city='г. Санкт-Петербург',
               tel='8 921 555-01-22', mail='Buh@KamenniyMost.example',
               dir='Белоусов Игорь Станиславович', dir_gen='Белоусова Игоря Станиславовича',
               dir_ini='Белоусов И.С.', dir_pos='директора'),
}
IP = dict(fio='Терехов Андрей Львович', fio_gen='Терехова Андрея Львовича', ini='Терехов А.Л.',
          inn=inn12('7804123456'), ogrnip=ogrnip('31578460012345'), bank='alfa',
          rs=rs('40802810', '044525593', '10000011923'), addr='ул. Рубинштейна, д. 9, кв. 31',
          city='г. Санкт-Петербург', tel='+7-911-222-33-44', mail='terekhov.al@mail.example')

def bank(k): return BANKS[k][0]

# ---------------------------------------------------------------- docx
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

def runs(text, bold=False):
    """Знак | в тексте режет абзац на несколько w:r с разным форматированием (обычный, жирный, курсив)."""
    out = []
    for i, seg in enumerate(text.split('|')):
        if not seg: continue
        pr = ''
        if bold or i % 3 == 1: pr = '<w:b/>'
        elif i % 3 == 2: pr = '<w:i/>'
        out.append(f'<w:r>{"<w:rPr>"+pr+"</w:rPr>" if pr else ""}<w:t xml:space="preserve">{escape(seg)}</w:t></w:r>')
    return ''.join(out)

def P(text, bold=False, center=False, size=None):
    ppr = ''
    if center: ppr += '<w:jc w:val="center"/>'
    if size: ppr += f'<w:rPr><w:sz w:val="{size}"/></w:rPr>'
    r = runs(text, bold)
    if size: r = r.replace('<w:r>', f'<w:r><w:rPr><w:sz w:val="{size}"/></w:rPr>').replace('<w:rPr><w:b/></w:rPr>', '<w:b/>') if False else r
    return f'<w:p>{"<w:pPr>"+ppr+"</w:pPr>" if ppr else ""}{r}</w:p>'

def H(text): return P(text, bold=True)

def TBL(rows, widths):
    x = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>' + ''.join(
        f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')) + '</w:tblBorders></w:tblPr>'
    for ri, row in enumerate(rows):
        x += '<w:tr>'
        for c, w in zip(row, widths):
            paras = c if isinstance(c, list) else [c]
            x += f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>' + ''.join(P(p, bold=(ri == 0 and len(rows) > 1 and False)) for p in paras) + '</w:tc>'
        x += '</w:tr>'
    return x + '</w:tbl>'

def docx(path, body, header, footer_text='Страница '):
    hdr = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W}><w:p>{runs(header)}</w:p></w:hdr>'
    ftr = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {W}><w:p>{runs(footer_text)}'
           '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>')
    doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>' + ''.join(body) +
           '<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="default" r:id="rId2"/>'
           '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
          '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels)
        z.writestr('word/_rels/document.xml.rels', drels); z.writestr('word/document.xml', doc)
        z.writestr('word/header1.xml', hdr); z.writestr('word/footer1.xml', ftr)

# ---------------------------------------------------------------- fods -> xlsx
FODS_HEAD = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" office:version="1.2" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
<office:automatic-styles>
<number:number-style style:name="N2"><number:number number:decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:number-style>
<number:date-style style:name="D1"><number:day number:style="long"/><number:text>.</number:text><number:month number:style="long"/><number:text>.</number:text><number:year number:style="long"/></number:date-style>
<style:style style:name="b" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style>
<style:style style:name="bm" style:family="table-cell" style:data-style-name="N2"><style:text-properties fo:font-weight="bold"/></style:style>
<style:style style:name="m" style:family="table-cell" style:data-style-name="N2"/>
<style:style style:name="d" style:family="table-cell" style:data-style-name="D1"/>
<style:style style:name="i1" style:family="table-cell"><style:paragraph-properties fo:margin-left="0.4cm"/></style:style>
<style:style style:name="i2" style:family="table-cell"><style:paragraph-properties fo:margin-left="0.8cm"/></style:style>
<style:style style:name="im" style:family="table-cell" style:data-style-name="N2"><style:paragraph-properties fo:margin-left="0.4cm"/></style:style>
</office:automatic-styles>
<office:body><office:spreadsheet>
'''

def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s

def cell(v, st=None):
    """v: None, str, int/float, ('date','2025-03-17'), ('f', формула, значение)"""
    a = f' table:style-name="{st}"' if st else ''
    if v is None: return f'<table:table-cell{a}/>'
    if isinstance(v, tuple) and v[0] == 'date':
        return f'<table:table-cell{a or " table:style-name=\"d\""} office:value-type="date" office:date-value="{v[1]}"><text:p>{v[1][8:]}.{v[1][5:7]}.{v[1][:4]}</text:p></table:table-cell>'
    if isinstance(v, tuple) and v[0] == 'f':
        return f'<table:table-cell{a or " table:style-name=\"m\""} table:formula="of:={v[1]}" office:value-type="float" office:value="{v[2]}"><text:p>{v[2]}</text:p></table:table-cell>'
    if isinstance(v, (int, float)):
        return f'<table:table-cell{a} office:value-type="float" office:value="{v}"><text:p>{v}</text:p></table:table-cell>'
    ps = ''.join(f'<text:p>{escape(l)}</text:p>' for l in str(v).split('\n'))
    return f'<table:table-cell{a} office:value-type="string">{ps}</table:table-cell>'

def sheet(name, rows):
    """rows: список строк, строка - список ячеек; ячейка v или (v, стиль)."""
    x = f'<table:table table:name="{escape(name)}">'
    for r in rows:
        x += '<table:table-row>'
        for c in r:
            if isinstance(c, tuple) and len(c) == 2 and isinstance(c[1], str) and c[0] is not None and c[0] != 'date' and c[0] != 'f' and c[1] in ('b', 'bm', 'm', 'd', 'i1', 'i2', 'im'):
                x += cell(c[0], c[1])
            elif isinstance(c, tuple) and len(c) == 2 and c[0] is None:
                x += cell(None, c[1])
            else:
                x += cell(c)
        x += '</table:table-row>'
    return x + '</table:table>'

def xlsx(path, sheets):
    d = tempfile.mkdtemp()
    f = os.path.join(d, 'book.fods')
    with open(f, 'w', encoding='utf-8') as fh:
        fh.write(FODS_HEAD + ''.join(sheets) + '</office:spreadsheet></office:body></office:document>')
    subprocess.run(['soffice', '--headless', '--convert-to', 'xlsx', '--outdir', d, f], check=True, capture_output=True)
    shutil.move(os.path.join(d, 'book.xlsx'), path)
    shutil.rmtree(d)

def money(v): return f'{v:,.2f}'.replace(',', ' ').replace('.', ',')

# ---------------------------------------------------------------- сделка 1
SV, KM = FIRMS['sv'], FIRMS['km']
N1, D1 = '41-КВ/2025', '17.03.2025'
PRICE1 = 4896000.00   # равно итогу акта № 7 (проверяется ниже)
ADV1 = round(PRICE1 * 0.3, 2)
MARKS = {}

def requisites(f, role):
    b = BANKS[f['bank']]
    return [role + ' ' + f['short'],
            f'ИНН {f["inn"]}, КПП {f["kpp"]}', f'ОГРН {f["ogrn"]}',
            f'р/с {f["rs"]}', f'в {b[0]}', f'к/с {b[2]}', f'БИК {b[1]}',
            f'Адрес: {b[3]}, {f["addr"]}', f'Тел.: {f["tel"]}', f'Эл. почта: {f["mail"]}']

def make_contract1():
    B = []
    B.append(P(f'ДОГОВОР ПОДРЯДА № {N1}', bold=True, center=True))
    B.append(P(f'г. Санкт-Петербург|{" " * 60}|{D1}'))
    B.append(P(f'Общество с ограниченной ответственностью «Каменный| Мос|т», именуемое в дальнейшем «Подрядчик», в лице {KM["dir_pos"]} {KM["dir_gen"]}, '
               f'действующего на основании Устава, с одной стороны, и Общество с ограниченной ответственностью «Северный Ветер| Строй|», именуемое в дальнейшем '
               f'«Заказчик», в лице {SV["dir_pos"]} {SV["dir_gen"]}, действующей на основании Устава, с другой стороны, совместно именуемые «Стороны», '
               'заключили настоящий договор о нижеследующем.'))
    B.append(H('1. Предмет договора'))
    B.append(P('1.1. Подрядчик обязуется выполнить по заданию Заказчика строительно-монтажные работы по устройству монолитного перекрытия и '
               'кирпичной кладке наружных стен на объекте «Складской комплекс, второй этап», а Заказчик обязуется принять результат работ и оплатить его.'))
    B.append(P(f'1.2. Объект расположен по адресу: г. Всеволожск, Ленинградская область, Колтушское шоссе, д. 44. Состав и объёмы работ определены сметой, '
               f'являющейся приложением к договору {N1}.'))
    B.append(P('1.3. Работы выполняются из материалов Подрядчика, если иное не оговорено в смете. Заказчик передаёт Подрядчику строительную площадку по акту.'))
    B.append(H('2. Цена договора и порядок расчётов'))
    B.append(P(f'2.1. Цена работ по настоящему договору составляет {money(PRICE1)} руб., в том числе НДС 20 %. Цена является твёрдой и пересмотру не подлежит, '
               'за исключением случаев, прямо предусмотренных договором.'))
    B.append(P(f'2.2. Заказчик перечисляет аванс в размере 30 % цены, то есть {money(ADV1)} руб., в течение пяти банковских дней с даты подписания договора на '
               f'расчётный счёт Подрядчика, указанный в разделе 9. Остаток оплачивается после подписания акта сдачи-приёмки.'))
    B.append(P('2.3. Обязательство по оплате считается исполненным в момент списания денежных средств с расчётного счёта Заказчика.'))
    B.append(H('3. Сроки выполнения работ'))
    B.append(P('3.1. Начало работ: 01.04.2025. Окончание работ: 30.06.2025. Промежуточные сроки определены графиком производства работ.'))
    B.append(P('3.2. При просрочке по вине Заказчика, включая несвоевременную передачу площадки, сроки продлеваются на соответствующий период.'))
    B.append(H('4. Приёмка работ'))
    B.append(P('4.1. По окончании работ Подрядчик направляет Заказчику акт сдачи-приёмки. Заказчик в течение пяти рабочих дней рассматривает акт и подписывает его '
               'либо направляет мотивированный отказ.'))
    B.append(P('4.2. Скрытые работы принимаются до их закрытия, о готовности к приёмке Подрядчик извещает Заказчика письменно.'))
    B.append(H('5. Гарантии'))
    B.append(P('5.1. Гарантийный срок на результат работ составляет 24 месяца с даты подписания акта сдачи-приёмки. Подрядчик за свой счёт устраняет дефекты, '
               'выявленные в течение гарантийного срока.'))
    B.append(H('6. Ответственность сторон'))
    B.append(P('6.1. За просрочку оплаты Заказчик уплачивает Подрядчику неустойку в размере 0,1 % от неоплаченной суммы за каждый день просрочки, но не более 10 % цены.'))
    B.append(P('6.2. За нарушение сроков выполнения работ Подрядчик уплачивает Заказчику неустойку в том же размере. Уплата неустойки не освобождает Сторону от исполнения обязательств.'))
    B.append(H('7. Конфиденциальность и разрешение споров'))
    B.append(P('7.1. Стороны не раскрывают третьим лицам условия договора и сведения, полученные при его исполнении, кроме случаев, предусмотренных законом.'))
    B.append(P('7.2. Споры решаются путём переговоров, срок ответа на претензию составляет 10 календарных дней. При недостижении согласия спор передаётся в '
               'Арбитражный суд города Санкт-Петербурга и Ленинградской области.'))
    B.append(H('8. Заключительные положения'))
    B.append(P('8.1. Договор вступает в силу с даты подписания и действует до полного исполнения Сторонами обязательств. Изменения оформляются письменным соглашением.'))
    B.append(P(f'8.2. Договор составлен в двух экземплярах. Ответственный за исполнение со стороны Заказчика - {SV["dir_ini"]}, со стороны Подрядчика - '
               f'главный инженер Лаптев Максим Олегович, тел. +7 (981) 700-45-19, эл. почта m.laptev@kamenniymost.example.'))
    B.append(H('9. Адреса, реквизиты и подписи сторон'))
    left = requisites(KM, 'Подрядчик:')
    right = requisites(SV, 'Заказчик:')
    left[0] = 'Подрядчик: |ООО «Каме|нный Мост|»'
    right[0] = 'Заказчик: ООО «Северный |Ветер| Строй»'
    left += [f'{KM["dir_pos"].capitalize()} ______________ / {KM["dir_ini"]} /']
    right += [f'Генеральный директор ______________ / {SV["dir_ini"]} /']
    B.append(TBL([['Подрядчик', 'Заказчик'], [left, right]], [4800, 4800]))
    docx(os.path.join(HERE, 'dogovor-podryada.docx'), B, f'Договор подряда № {N1}, конфиденциально')
    b1, b2 = BANKS['vtb'], BANKS['sber']
    MARKS['dogovor-podryada.docx'] = dict(
        must_types={
            KM['core']: 'организация', SV['core']: 'организация',
            KM['dir_gen']: 'фио', SV['dir_gen']: 'фио', SV['dir_ini']: 'фио', KM['dir_ini']: 'фио', 'Лаптев Максим Олегович': 'фио',
            KM['inn']: 'инн', SV['inn']: 'инн', KM['rs']: 'счёт', SV['rs']: 'счёт', b1[2]: 'счёт', b2[2]: 'счёт',
            KM['tel']: 'телефон', SV['tel']: 'телефон', '+7 (981) 700-45-19': 'телефон',
            KM['mail']: 'почта', SV['mail']: 'почта', 'm.laptev@kamenniymost.example': 'почта',
            KM['addr']: 'адрес', SV['addr']: 'адрес', 'Колтушское шоссе, д. 44': 'адрес',
            N1: 'номер_договора', money(PRICE1): 'сумма', money(ADV1): 'сумма',
            D1: 'дата', KM['ogrn']: 'прочее', SV['ogrn']: 'прочее', KM['kpp']: 'прочее', SV['kpp']: 'прочее'},
        keep=['Предмет договора', 'Цена договора и порядок расчётов', 'Сроки выполнения работ', 'Приёмка работ', 'Гарантии',
              'Ответственность сторон', 'Адреса, реквизиты и подписи сторон', 'строительно-монтажные работы', 'монолитного перекрытия',
              'Арбитражный суд города Санкт-Петербурга и Ленинградской области', b1[0], b2[0], 'г. Всеволожск', 'главный инженер',
              'Подрядчик', 'Заказчик', 'Генеральный директор', 'акт сдачи-приёмки', 'Складской комплекс, второй этап'])

ACT_NO, ACT_DATE = '7', '30.06.2025'
WORKS = [('Устройство монолитного перекрытия, толщина 200 мм', 'м2', 1150, 1960.00),
         ('Кирпичная кладка наружных стен', 'м3', 380, 4200.00),
         ('Устройство армопояса', 'м.п.', 520, 1150.00),
         ('Гидроизоляция фундаментных стен', 'м2', 640, 700.00)]
WSUM = [round(q * p, 2) for _, _, q, p in WORKS]
ACT_TOTAL = round(sum(WSUM), 2)
assert ACT_TOTAL == PRICE1, ACT_TOTAL

def make_act1():
    B = [P(f'АКТ № {ACT_NO} СДАЧИ-ПРИЁМКИ ВЫПОЛНЕННЫХ РАБОТ', bold=True, center=True),
         P(f'г. Санкт-Петербург|{" " * 60}|{ACT_DATE}'),
         P(f'Мы, нижеподписавшиеся, {KM["dir"].replace("Белоусов", "Белоусов")}, {KM["dir_pos"]} {KM["short"]} (Подрядчик), с одной стороны, и '
           f'{SV["dir_pos"]} {SV["short"]} {SV["dir"]} (Заказчик), с другой стороны, составили настоящий акт о том, что Подрядчик выполнил, '
           f'а Заказчик принял работы по договору подряда № {N1} от {D1}.')]
    rows = [['№', 'Наименование работ', 'Ед. изм.', 'Количество', 'Цена, руб.', 'Сумма, руб.']]
    for i, ((n, u, q, p), s) in enumerate(zip(WORKS, WSUM), 1):
        rows.append([str(i), n, u, str(q), money(p), money(s)])
    rows.append(['', 'Итого', '', '', '', money(ACT_TOTAL)])
    B.append(TBL(rows, [600, 3900, 900, 1300, 1400, 1500]))
    B.append(P(f'Всего выполнено работ на сумму {money(ACT_TOTAL)} руб., в том числе НДС 20 %. Работы выполнены полностью и в срок, Заказчик претензий по объёму, '
               'качеству и срокам не имеет.'))
    B.append(P(f'Ранее перечислен аванс {money(ADV1)} руб., к доплате по настоящему акту {money(ACT_TOTAL - ADV1)} руб. Оплата производится на расчётный счёт '
               f'Подрядчика {KM["rs"]} в {bank(KM["bank"])}.'))
    B.append(TBL([['От Подрядчика', 'От Заказчика'],
                  [[KM['short'], f'{KM["dir_pos"].capitalize()} ____________ {KM["dir_ini"]}', f'ИНН {KM["inn"]}'],
                   ['Северный Ветер Строй ООО', f'Генеральный директор ____________ {SV["dir_ini"]}', f'ИНН {SV["inn"]}']]], [4800, 4800]))
    docx(os.path.join(HERE, 'akt-sdachi-priemki.docx'), B, f'Акт № {ACT_NO} к договору № {N1} от {D1}')
    MARKS['akt-sdachi-priemki.docx'] = dict(
        must_types={KM['core']: 'организация', SV['core']: 'организация', KM['dir']: 'фио', SV['dir']: 'фио', KM['dir_ini']: 'фио', SV['dir_ini']: 'фио',
                    KM['inn']: 'инн', SV['inn']: 'инн', KM['rs']: 'счёт', N1: 'номер_договора', money(ACT_TOTAL): 'сумма',
                    money(ADV1): 'сумма', money(ACT_TOTAL - ADV1): 'сумма', D1: 'дата', ACT_DATE: 'дата'},
        keep=['АКТ', 'СДАЧИ-ПРИЁМКИ ВЫПОЛНЕННЫХ РАБОТ', 'Наименование работ', 'Ед. изм.', 'Количество', 'Итого', 'Подрядчик', 'Заказчик',
              'Устройство монолитного перекрытия', 'Кирпичная кладка наружных стен', 'Устройство армопояса', 'Гидроизоляция фундаментных стен',
              'Генеральный директор', bank(KM['bank'])])

def make_sverka1():
    L = 'Дебет'; R = 'Кредит'
    docs = [  # документ, дата, сумма, K - наше поступление (кредит 60), D - наша оплата (дебет 60)
        (f'Поступление (акт, накладная) № {ACT_NO}', ACT_DATE, ACT_TOTAL, 'K'),
        ('Платёжное поручение (исходящее) № 214', '24.03.2025', ADV1, 'D'),
        ('Платёжное поручение (исходящее) № 398', '18.04.2025', 1500000.00, 'D'),
        ('Платёжное поручение (исходящее) № 561', '03.06.2025', 1000000.00, 'D')]
    r = []
    r.append([('Акт сверки взаимных расчётов', 'b')])
    r.append(['за период: 01.01.2025 - 30.06.2025'])
    r.append([f'между {SV["short"]} и {KM["short"]}'])
    r.append([f'по договору № {N1} от {D1}'])
    r.append([f'Мы, нижеподписавшиеся, {SV["dir_pos"]} {SV["short"]} {SV["dir_ini"]}, с одной стороны, и {KM["dir_pos"]} {KM["short"]} {KM["dir_ini"]}, '
              'с другой стороны, составили настоящий акт сверки в том, что состояние взаимных расчётов по данным учёта следующее:'])
    r.append([])
    r.append([(f'По данным {SV["short"]}, руб.', 'b'), None, None, None, (f'По данным {KM["short"]}, руб.', 'b'), None, None])
    r.append([('Содержание записи', 'b'), (L, 'b'), (R, 'b'), None, ('Содержание записи', 'b'), (L, 'b'), (R, 'b')])
    r.append(['Сальдо начальное', None, None, None, 'Сальдо начальное', None, None])
    first = len(r) + 1
    for n, d, s, side in docs:
        left = f'{n} от {d}'
        right = left.replace('Поступление (акт, накладная)', 'Реализация (акт, накладная)').replace('Платёжное поручение (исходящее)', 'Платёжное поручение (входящее)')
        if side == 'K':   # у нас кредит 60, у подрядчика дебет 62
            r.append([left, None, s, None, right, s, None])
        else:
            r.append([left, s, None, None, right, None, s])
    last = len(r)
    sd = sum(x[2] for x in docs if x[3] == 'D'); sk = sum(x[2] for x in docs if x[3] == 'K')
    def fsum(c, v): return ('f', f'SUM([.{c}{first}:.{c}{last}])', v)
    r.append([('Обороты за период', 'b'), (fsum('B', sd), 'bm'), (fsum('C', sk), 'bm'), None, ('Обороты за период', 'b'), (fsum('F', sk), 'bm'), (fsum('G', sd), 'bm')])
    tr = len(r)
    end = sk - sd
    r.append([('Сальдо конечное', 'b'), None, (('f', f'C{tr}-B{tr}', end), 'bm'), None, ('Сальдо конечное', 'b'), (('f', f'F{tr}-G{tr}', end), 'bm'), None])
    r.append([f'по данным {SV["short"]} на 30.06.2025 задолженность в пользу {KM["short"]} {money(end)} руб.' if end else
              f'по данным {SV["short"]} на 30.06.2025 задолженности нет'])
    r.append([])
    r.append([f'От {SV["short"]}', None, None, None, f'От {KM["short"]}'])
    r.append([f'{SV["dir_pos"].capitalize()} ____________ {SV["dir_ini"]}', None, None, None, f'{KM["dir_pos"].capitalize()} ____________ {KM["dir_ini"]}'])
    r.append(['М.П.', None, None, None, 'М.П.'])
    xlsx(os.path.join(HERE, 'akt-sverki-1c.xlsx'), [sheet('Акт сверки', r)])
    MARKS['akt-sverki-1c.xlsx'] = dict(
        must_types={SV['core']: 'организация', KM['core']: 'организация', SV['dir_ini']: 'фио', KM['dir_ini']: 'фио', N1: 'номер_договора', D1: 'дата'},
        keep=['Акт сверки взаимных расчётов', 'Содержание записи', 'Дебет', 'Кредит', 'Сальдо начальное', 'Сальдо конечное', 'Обороты за период',
              'Поступление (акт, накладная)', 'Реализация (акт, накладная)', 'Платёжное поручение (исходящее)',
              str(int(ADV1)), '1500000', str(int(sk)), 'М.П.', 'Генерального директора'])
    return end

# ---------------------------------------------------------------- сделка 2
CONTR = [  # ключ, вид, название, ИНН (десятизначный или двенадцатизначный)
    ('km', 'ООО', 'Каменный Мост', KM['inn']),
    ('ip', 'ИП', 'Терехов А.Л.', IP['inn']),
    ('c3', 'АО', 'Нева-Металл', inn10('780512340')),
    ('c4', 'ООО', 'Техснаб-Урал', inn10('665811223')),
    ('c5', 'ООО', 'Пионер Сервис', inn10('780244561')),
    ('c6', 'АО', 'Балтийская Оптовая Компания', inn10('780377102')),
    ('c7', 'ИП', 'Хамидуллина Р.Ф.', inn12('1660987123')),
    ('c8', 'ООО', 'Лесной Двор', inn10('471011890')),
    ('c9', 'ООО', 'Горизонт Монтаж', inn10('781405672')),
    ('c10', 'ИП', 'Орлов Е.В.', inn12('4703112233')),
    ('c11', 'АО', 'Вега Транс', inn10('780699341')),
    ('c12', 'ООО', 'Радуга Опт', inn10('781138805')),
    ('c13', 'ООО', 'Стальной Дом', inn10('780911764')),
    ('c14', 'ИП', 'Макарова Т.С.', inn12('7810556677')),
]
def cname(c):
    return f'ИП {c[2]}' if c[1] == 'ИП' else f'{c[1]} «{c[2]}»'

# договоры по контрагентам: (номер, дата ГГГГ-ММ-ДД, предмет, сумма, срок)
CONTRACTS = {
    'km': [('41-КВ/2025', '2025-03-17', 'Подряд на строительно-монтажные работы', PRICE1, '2025-06-30')],
    'ip': [('12/03-П', '2025-03-12', 'Поставка крепежа и метизов', 864300.50, '2025-12-31')],
    'c3': [('ДП-0045', '2024-09-02', 'Поставка арматуры', 5120000.00, '2025-09-02'), ('ДП-0061', '2025-02-10', 'Поставка листового проката', 2210000.00, '2025-12-31')],
    'c4': [('ТУ-118/24', '2024-11-20', 'Аренда башенного крана', 960000.00, '2025-05-20')],
    'c5': [('ПС-07/25', '2025-01-15', 'Уборка объектов', 312000.00, '2025-12-31')],
    'c6': [('ДП-0052', '2024-12-05', 'Поставка бетона и раствора', 4075000.00, '2025-06-05')],
    'c7': [('Х-03/2025', '2025-02-01', 'Услуги автокрана', 187500.00, '2025-07-31')],
    'c8': [('ЛД-21', '2024-10-14', 'Поставка пиломатериалов', 698400.00, '2025-04-14')],
    'c9': [('ГМ-0310', '2025-03-10', 'Монтаж инженерных сетей', 1740000.00, '2025-08-31')],
    'c10': [('О/2025-8', '2025-04-01', 'Геодезические изыскания', 129000.00, '2025-06-01')],
    'c11': [('ВТ-45/2025', '2025-01-20', 'Перевозка грузов', 455000.00, '2025-12-31')],
    'c12': [('РО-0099', '2024-08-19', 'Поставка кирпича', 2630000.00, '2025-03-31')],
    'c13': [('СД-14/25', '2025-02-24', 'Поставка металлоконструкций', 3390000.00, '2025-10-31')],
    'c14': [('М-12/05', '2025-05-12', 'Услуги технического надзора', 96000.00, '2025-12-31')],
}
RESP = ['Сидоров Павел Константинович', 'Гусева Елена Андреевна', 'Фомин Д.Р.', 'Нестерова Алла Викторовна']
STAFF = [  # ФИО, таб., должность, оклад, дата приёма, телефон
    ('Нестерова Алла Викторовна', '0001', 'Генеральный директор', 210000, '2012-04-02', '+7 (812) 346-18-02'),
    ('Сидоров Павел Константинович', '0007', 'Главный инженер', 165000, '2014-08-18', '8 911 903-27-45'),
    ('Гусева Елена Андреевна', '0012', 'Главный бухгалтер', 150000, '2015-02-09', '+7-921-318-64-70'),
    ('Фомин Дмитрий Романович', '0019', 'Начальник отдела снабжения', 120000, '2017-06-26', '89817004519'),
    ('Ковалёва Ирина Петровна', '0023', 'Бухгалтер', 78000, '2018-09-03', '+7 (812) 346-18-05'),
    ('Прохоров Артём Игоревич', '0031', 'Прораб', 112000, '2019-03-11', '8 (921) 555-30-18'),
    ('Яковлева Наталья Сергеевна', '0035', 'Юрист', 98000, '2020-01-13', '+7 905 211-76-39'),
    ('Мартынов Олег Викторович', '0042', 'Менеджер по закупкам', 85000, '2021-05-24', '+7-952-640-12-88'),
    ('Зайцева Мария Олеговна', '0046', 'Секретарь', 54000, '2022-10-17', '8 812 346 18 00'),
    ('Дроздов Кирилл Алексеевич', '0051', 'Инженер ПТО', 92000, '2023-07-03', '89112847706'),
]
STAFF_FIO = [s[0] for s in STAFF]

def make_contract2():
    NUM, DT = '12/03-П', '12.03.2025'
    SPEC = [('Болт анкерный М12х120', 'шт.', 4000, 38.50), ('Саморез кровельный 5,5х38', 'шт.', 20000, 4.90),
            ('Шпилька резьбовая М16, 1 м', 'шт.', 800, 95.00), ('Гайка шестигранная М16', 'шт.', 3200, 7.80),
            ('Уголок монтажный 50х50', 'шт.', 1500, 54.00)]
    ssum = [round(q * p, 2) for _, _, q, p in SPEC]; total = round(sum(ssum), 2)
    B = [P(f'ДОГОВОР ПОСТАВКИ № {NUM}', bold=True, center=True), P(f'г. Санкт-Петербург|{" " * 60}|{DT}'),
         P(f'Индивидуальный предприниматель |Терехов Андрей Львович|, ОГРНИП {IP["ogrnip"]}, ИНН {IP["inn"]}, именуемый в дальнейшем «Поставщик», '
           f'с одной стороны, и {SV["short"]}, именуемое в дальнейшем «Покупатель», в лице {SV["dir_pos"]} {SV["dir_gen"]}, действующей на основании Устава, '
           'с другой стороны, заключили настоящий договор о нижеследующем.'),
         H('1. Предмет договора'),
         P('1.1. Поставщик обязуется передавать в собственность Покупателя крепёжные изделия и метизы (далее - товар), а Покупатель обязуется принимать товар и оплачивать его. '
           'Ассортимент, количество и цена определяются спецификацией.'),
         P(f'1.2. Спецификация является неотъемлемой частью договора {NUM}. Товар поставляется партиями по заявкам Покупателя.'),
         H('2. Цена и расчёты'),
         P(f'2.1. Общая стоимость товара по спецификации составляет {money(total)} руб. НДС не облагается: Поставщик применяет упрощённую систему налогообложения.'),
         P(f'2.2. Оплата производится в течение семи календарных дней с даты поставки партии на расчётный счёт Поставщика {IP["rs"]} в {bank(IP["bank"])}.'),
         H('3. Поставка и приёмка'),
         P('3.1. Доставка осуществляется силами Поставщика на склад Покупателя по адресу: г. Санкт-Петербург, ул. Балканская, д. 20, склад 3. '
           'Риск случайной гибели товара переходит к Покупателю с момента подписания товарной накладной.'),
         P('3.2. Приёмка по количеству и качеству производится при получении. Претензии по недостаче и браку заявляются в течение трёх рабочих дней.'),
         H('4. Ответственность и срок действия'),
         P('4.1. За просрочку поставки Поставщик уплачивает неустойку 0,05 % стоимости непоставленного товара за каждый день. Договор действует по 31.12.2025.'),
         P(f'4.2. Контактное лицо Покупателя - {STAFF[3][0]}, тел. {STAFF[3][5]}. Контактное лицо Поставщика - {IP["ini"]}, тел. {IP["tel"]}, эл. почта {IP["mail"]}.'),
         H('5. Реквизиты сторон'),
         TBL([['Поставщик', 'Покупатель'],
              [[f'ИП {IP["ini"]}', f'ИНН {IP["inn"]}', f'ОГРНИП {IP["ogrnip"]}', f'р/с {IP["rs"]}', f'в {bank(IP["bank"])}', f'БИК {BANKS[IP["bank"]][1]}',
                f'Адрес: {BANKS[IP["bank"]][3]}, {IP["addr"]}'],
               [SV['short'], f'ИНН {SV["inn"]}, КПП {SV["kpp"]}', f'р/с {SV["rs"]}', f'в {bank(SV["bank"])}', f'БИК {BANKS[SV["bank"]][1]}', f'Тел.: {SV["tel"]}']]],
             [4800, 4800]),
         H('Спецификация № 1'),
         ]
    rows = [['№', 'Наименование', 'Ед. изм.', 'Кол-во', 'Цена, руб.', 'Сумма, руб.']]
    for i, ((n, u, q, p), s) in enumerate(zip(SPEC, ssum), 1): rows.append([str(i), n, u, str(q), money(p), money(s)])
    rows.append(['', 'Итого', '', '', '', money(total)])
    B.append(TBL(rows, [600, 3700, 900, 1000, 1500, 1900]))
    B.append(P(f'Поставщик: ______________ / {IP["ini"]} /      Покупатель: ______________ / {SV["dir_ini"]} /'))
    docx(os.path.join(HERE, 'dogovor-postavki-ip.docx'), B, f'Договор поставки № {NUM}')
    MARKS['dogovor-postavki-ip.docx'] = dict(
        must_types={IP['fio']: 'фио', IP['ini']: 'фио', SV['dir_gen']: 'фио', SV['dir_ini']: 'фио', STAFF[3][0]: 'фио', SV['core']: 'организация',
                    IP['inn']: 'инн', SV['inn']: 'инн', IP['rs']: 'счёт', SV['rs']: 'счёт', IP['tel']: 'телефон', STAFF[3][5]: 'телефон', SV['tel']: 'телефон',
                    IP['mail']: 'почта', IP['addr']: 'адрес', 'ул. Балканская, д. 20, склад 3': 'адрес', NUM: 'номер_договора',
                    money(total): 'сумма', money(ssum[0]): 'сумма', DT: 'дата', IP['ogrnip']: 'прочее', SV['kpp']: 'прочее'},
        keep=['Предмет договора', 'Цена и расчёты', 'Поставка и приёмка', 'Ответственность и срок действия', 'Реквизиты сторон', 'Спецификация № 1',
              'крепёжные изделия и метизы', 'Болт анкерный М12х120', 'Шпилька резьбовая М16', 'Уголок монтажный 50х50', 'Поставщик', 'Покупатель',
              'Индивидуальный предприниматель', bank(IP['bank']), bank(SV['bank']), 'Итого', 'Наименование', 'Ед. изм.'])

def build_osv():
    """Сальдо по контрагенту и договорам; возвращает строки листа и итоги."""
    r = []
    r.append([(f'Оборотно-сальдовая ведомость по счету 60', 'b')])
    r.append([SV['short']])
    r.append(['за 1 квартал 2025 г.'])
    r.append(['Счёт 60 - Расчёты с поставщиками и подрядчиками'])
    r.append([])
    r.append([('Контрагент / Договор', 'b'), ('Сальдо на начало периода', 'b'), None, ('Обороты за период', 'b'), None, ('Сальдо на конец периода', 'b'), None])
    r.append([None, ('Дебет', 'b'), ('Кредит', 'b'), ('Дебет', 'b'), ('Кредит', 'b'), ('Дебет', 'b'), ('Кредит', 'b')])
    rowno = len(r)
    contr_rows = []; groups = {}
    seed = 3
    for c in CONTR:
        start = rowno + 1
        rowno += 1
        r.append(None); idx = len(r) - 1
        kids = []
        for k in CONTRACTS[c[0]]:
            seed = (seed * 7 + 11) % 97
            nach_kt = round((seed * 5200.0) if seed % 3 else 0.0, 2)
            ob_dt = round(k[3] * (0.35 + (seed % 5) / 10), 2)
            ob_kt = round(k[3] * (0.4 + (seed % 4) / 10), 2)
            kids.append((k, 0.0, nach_kt, ob_dt, ob_kt))
            rowno += 1
        groups[c[0]] = (idx, start, kids)
        # строка контрагента с формулами сумм по договорам
        s = start + 1; e = start + len(kids)
        sums = [sum(x[i] for x in kids) for i in (1, 2, 3, 4)]
        vals = []
        for ci, v in zip('BCDE', sums): vals.append((('f', f'SUM([.{ci}{s}:.{ci}{e}])', round(v, 2)), 'bm'))
        net = (sums[1] - sums[0]) + (sums[3] - sums[2])
        edt, ekt = (0.0, round(net, 2)) if net >= 0 else (round(-net, 2), 0.0)
        row = [(cname(c), 'b')] + vals[:4]
        row.append((('f', f'MAX([.B{start}]-[.C{start}]+[.D{start}]-[.E{start}];0)', edt), 'bm'))
        row.append((('f', f'MAX([.C{start}]-[.B{start}]+[.E{start}]-[.D{start}];0)', ekt), 'bm'))
        r[idx] = row
        contr_rows.append(start)
        for k, a, b, dd, kk in kids:
            net = (b - a) + (kk - dd)
            edt, ekt = (0.0, round(net, 2)) if net >= 0 else (round(-net, 2), 0.0)
            r.append([(f'Договор № {k[0]} от {k[1][8:]}.{k[1][5:7]}.{k[1][:4]}', 'i1'), (a or None, 'im'), (b or None, 'im'), (dd, 'im'), (kk, 'im'),
                      (edt or None, 'im'), (ekt or None, 'im')])
    # итого по контрагентам
    tot = []
    for col_i, ci in enumerate('BCDEFG'):
        refs = '+'.join(f'[.{ci}{n}]' for n in contr_rows)
        val = 0.0
        for c in CONTR:
            row = r[groups[c[0]][0]]
            cv = row[col_i + 1]
            v = cv[0][2] if isinstance(cv, tuple) and isinstance(cv[0], tuple) else 0
            val += v
        tot.append((('f', refs, round(val, 2)), 'bm'))
    r.append([('Итого', 'b')] + tot)
    return r, groups

def make_osv():
    r, groups = build_osv()
    xlsx(os.path.join(HERE, 'osv-60.xlsx'), [sheet('ОСВ 60', r)])
    must = {}
    for c in CONTR:
        must[c[2].replace(' А.Л.', ' А.Л.')] = 'фио' if c[1] == 'ИП' else 'организация'
    ks = [f'Договор № {k[0]}' for c in CONTR for k in CONTRACTS[c[0]]]
    nums = {k[0]: 'номер_договора' for c in CONTR for k in CONTRACTS[c[0]]}
    must.update(nums)
    must[SV['core']] = 'организация'
    MARKS['osv-60.xlsx'] = dict(
        must_types=must,
        keep=['Оборотно-сальдовая ведомость по счету 60', 'за 1 квартал 2025 г.', 'Расчёты с поставщиками и подрядчиками', 'Контрагент / Договор',
              'Сальдо на начало периода', 'Обороты за период', 'Сальдо на конец периода', 'Дебет', 'Кредит', 'Итого'])

def make_card():
    c0 = CONTR[2]  # Нева-Металл, один контрагент с двумя договорами
    d1, d2 = CONTRACTS['c3']
    ent = [('2025-01-14', 'Поступление (акт, накладная) 0000012 от 14.01.2025', d1[0], 'K', 1820000.00),
           ('2025-01-23', 'Списание с расчётного счёта 00000006 от 23.01.2025', d1[0], 'D', 1500000.00),
           ('2025-02-12', 'Поступление (акт, накладная) 0000031 от 12.02.2025', d2[0], 'K', 1105000.00),
           ('2025-02-28', 'Списание с расчётного счёта 00000011 от 28.02.2025', d2[0], 'D', 600000.00),
           ('2025-03-19', 'Поступление (акт, накладная) 0000058 от 19.03.2025', d1[0], 'K', 950000.00),
           ('2025-03-27', 'Списание с расчётного счёта 00000019 от 27.03.2025', d1[0], 'D', 1200000.00)]
    r = [[('Карточка счета 60', 'b')], [f'{SV["short"]}'], ['за 1 квартал 2025 г.'], [f'Контрагент: {cname(c0)}, ИНН {c0[3]}'],
         [f'Договоры: № {d1[0]}, № {d2[0]}'], [],
         [('Период', 'b'), ('Документ', 'b'), ('Аналитика Дт', 'b'), ('Аналитика Кт', 'b'), ('Дебет', 'b'), ('Кредит', 'b'), ('Текущее сальдо', 'b')],
         [None, ('Сальдо на начало', 'b'), None, None, None, None, 'Кт 0,00']]
    first = len(r) + 1; bal = 0.0
    for d, doc, dog, side, s in ent:
        dt = ('date', d)
        an_tov = f'10.01\nАрматура и прокат\nДоговор № {dog}'
        an_k = f'{cname(c0)}\nДоговор № {dog}'
        if side == 'K':
            bal += s; r.append([(dt, 'd'), doc, an_tov, f'60.01\n{an_k}', None, (s, 'm'), f'Кт {money(bal)}'])
        else:
            bal -= s; r.append([(dt, 'd'), doc, f'60.01\n{an_k}', '51\nОсновной расчётный счёт', (s, 'm'), None, f'Кт {money(bal)}'])
    last = len(r)
    sd = sum(e[4] for e in ent if e[3] == 'D'); sk = sum(e[4] for e in ent if e[3] == 'K')
    r.append([None, ('Обороты за период и сальдо на конец', 'b'), None, None, (('f', f'SUM([.E{first}:.E{last}])', sd), 'bm'), (('f', f'SUM([.F{first}:.F{last}])', sk), 'bm'),
              (f'Кт {money(sk - sd)}', 'b')])
    xlsx(os.path.join(HERE, 'karta-60.xlsx'), [sheet('Карточка 60', r)])
    must = {'Нева-Металл': 'организация', c0[3]: 'инн', d1[0]: 'номер_договора', d2[0]: 'номер_договора', SV['core']: 'организация'}
    MARKS['karta-60.xlsx'] = dict(must_types=must,
        keep=['Карточка счета 60', 'за 1 квартал 2025 г.', 'Период', 'Документ', 'Аналитика Дт', 'Аналитика Кт', 'Текущее сальдо', 'Сальдо на начало',
              'Основной расчётный счёт', 'Арматура и прокат', 'Поступление (акт, накладная)', 'Списание с расчётного счёта', 'Обороты за период и сальдо на конец'])

def make_registry():
    def sheet_for(year):
        rows = [[('Реестр договоров ' + year, 'b')], [SV['short']], [],
                [('№ договора', 'b'), ('Дата', 'b'), ('Контрагент', 'b'), ('ИНН', 'b'), ('Предмет', 'b'), ('Сумма', 'b'), ('Срок', 'b'), ('Ответственный', 'b')]]
        i = 0
        for c in CONTR:
            for k in CONTRACTS[c[0]]:
                if k[1][:4] != year: continue
                resp = RESP[i % len(RESP)]; i += 1
                rows.append([k[0], (('date', k[1]), 'd'), cname(c) if i % 2 else (c[2] + (' ' + c[1] if c[1] != 'ИП' else '')) if c[1] != 'ИП' else cname(c),
                             int(c[3]), k[2], (k[3], 'm'), (('date', k[4]), 'd'), resp])
        return rows
    xlsx(os.path.join(HERE, 'reestr-dogovorov.xlsx'), [sheet('2024', sheet_for('2024')), sheet('2025', sheet_for('2025'))])
    must = {}
    for c in CONTR:
        must[str(int(c[3]))] = 'инн'
    for c in CONTR:
        for k in CONTRACTS[c[0]]:
            must[k[0]] = 'номер_договора'
    for n in ['Сидоров Павел Константинович', 'Гусева Елена Андреевна', 'Фомин Д.Р.', 'Нестерова Алла Викторовна']: must[n] = 'фио'
    for c in CONTR: must[c[2]] = 'фио' if c[1] == 'ИП' else 'организация'
    MARKS['reestr-dogovorov.xlsx'] = dict(must_types=must,
        keep=['Реестр договоров 2024', 'Реестр договоров 2025', '№ договора', 'Контрагент', 'Предмет', 'Сумма', 'Срок', 'Ответственный',
              'Поставка арматуры', 'Аренда башенного крана', 'Монтаж инженерных сетей', 'Услуги технического надзора'])

def make_staff():
    r = [[('Штатное расписание', 'b')], [SV['short']], ['по состоянию на 01.10.2025'], [],
         [('ФИО', 'b'), ('Табельный номер', 'b'), ('Должность', 'b'), ('Оклад', 'b'), ('Дата приёма', 'b'), ('Телефон', 'b')]]
    for f, t, p, o, d, ph in STAFF:
        r.append([f, t, p, (o, 'm'), (('date', d), 'd'), ph])
    f0 = 6; f1 = len(r)
    r.append([('Итого фонд оплаты труда', 'b'), None, None, (('f', f'SUM([.D{f0}:.D{f1}])', sum(s[3] for s in STAFF)), 'bm')])
    xlsx(os.path.join(HERE, 'shtatnoe-raspisanie.xlsx'), [sheet('Штатное расписание', r)])
    must = {f: 'фио' for f in STAFF_FIO}
    for s in STAFF: must[s[5]] = 'телефон'
    must[SV['core']] = 'организация'
    MARKS['shtatnoe-raspisanie.xlsx'] = dict(must_types=must,
        keep=['Штатное расписание', 'Табельный номер', 'Должность', 'Оклад', 'Дата приёма', 'Телефон', 'Генеральный директор', 'Главный инженер', 'Главный бухгалтер',
              'Начальник отдела снабжения', 'Менеджер по закупкам', 'Итого фонд оплаты труда'])

# ---------------------------------------------------------------- main
if __name__ == '__main__':
    make_contract1(); make_act1(); make_sverka1()
    make_contract2(); make_osv(); make_card(); make_registry(); make_staff()
    order = ['dogovor-podryada.docx', 'akt-sdachi-priemki.docx', 'akt-sverki-1c.xlsx', 'dogovor-postavki-ip.docx',
             'osv-60.xlsx', 'karta-60.xlsx', 'reestr-dogovorov.xlsx', 'shtatnoe-raspisanie.xlsx']
    out = [dict(file=f, must=[], must_types=MARKS[f]['must_types'], keep=MARKS[f]['keep']) for f in order]
    with open(os.path.join(ROOT, 'tests', 'deal.json'), 'w', encoding='utf-8') as fh:
        json.dump(out, fh, ensure_ascii=False, indent=1)
        fh.write('\n')
    print('готово:', ', '.join(order))
