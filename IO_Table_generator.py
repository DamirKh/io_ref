#!/usr/bin/env python
# -*- coding: utf-8 -*-
import csv
import re
import sys
import datetime
import argparse
from pathlib import Path
import xlsxwriter
import l5x
import re

io_config = {}
io_description = {}


class Tag(object):
    Other, DI, DO, AI, AO, = 0, 1, 2, 3, 4
    types = (Other, DI, DO, AI, AO,)

    def __init__(self, name: str, type=0):
        self.name = name
        self._type = type

    @property
    def type(self):
        return self._type

    @type.setter
    def type(self, value):
        if value in Tag.types:
            self._type = value

    @property
    def kip_name(self) -> str:
        kip = self.name.removeprefix('i').removeprefix('o')
        match = re.match(r"([A-Z]+)([0-9]+[A-Z]*)", kip, re.IGNORECASE)
        if match:
            head, tail = match.groups()
            out = f'{head}-{tail}'
            return out
        else:
            return kip

    def __str__(self):
        return self.name


def RUS_comment_decoder(comment: str):
    """ Decode russian comments"""
    # $0422$0435$043a$0443$0449$0430$044f $0441$0442$0435$043f$0435$043d$044c $043e$0442$043a$0440$044b$0442$0438$044f, %
    # Текущая степень открытия, %
    if comment is None:
        return ""
    out = ''
    pos = 0
    try:
        while pos < len(comment):
            if comment[pos] == '$':
                if comment[pos + 1] == 'Q' or comment[pos + 1] == 'N':
                    out += '\n'
                    pos += 2
                    continue
                rus_symbol_code = comment[pos + 1:pos + 5]
                out += chr(int(rus_symbol_code, base=16))
                pos += 5
            else:
                out += comment[pos]
                pos += 1
    except IndexError:
        pass
    return out


def tag2kip(tag_name: str):  # TODO convert to class method
    kip = tag_name.removeprefix('i').removeprefix('o')
    if '_' in kip:
        return kip
    match = re.match(r"([A-Z]+)([0-9]+[A-Z]*)", kip, re.IGNORECASE)
    if match:
        head, tail = match.groups()
        out = f'{head}-{tail}'
        return out
    else:
        return kip


def append_chass(chass_name: str, slot_num: int):
    global io_config
    global io_description
    if chass_name not in io_config.keys():
        io_config[chass_name] = {}
    if slot_num not in io_config[chass_name]:
        io_config[chass_name][slot_num] = {}
    # bad copy paste
    if chass_name not in io_description.keys():
        io_description[chass_name] = {}
    if slot_num not in io_description[chass_name]:
        io_description[chass_name][slot_num] = {}


class n11mapping(object):
    def __init__(self, map_file_name):
        with open(map_file_name, newline='') as map_file:
            map_reader = csv.reader(map_file, delimiter=' ')
            self._n11 = {}
            try:
                for row in map_reader:
                    # N11[0] CP_P0024JA:6:I.Data
                    if row[0].startswith('#'):
                        continue
                    self._n11[row[0]] = row[1]
                    # CP_P0024JA:6:I.Data N11[0]
            except IndexError:
                print(f"Unparsed row in map file '{map_file_name}'")
            print(f'Read {len(self._n11.keys())} point from map file')

    def replace(self, point_address: str):
        for n in self._n11.keys():  # N11[0].
            if point_address.startswith(n):
                return point_address.replace(n, self._n11[n])
        return point_address


def read_input_csv(filename, map_file_name=None, old_csv_version=False):
    global io_config
    global io_description

    print(f'Input file name = "{filename}"')
    if map_file_name:
        print(f'Map file name = "{map_file_name}"')
        n11 = n11mapping(map_file_name)
        map_func = n11.replace
    else:
        map_func = lambda s: s

    csv_delimiter = '?' if old_csv_version else ','

    with open(filename, newline='', encoding="ISO-8859-1") as csvfile:
        spamreader = csv.reader(csvfile, delimiter=csv_delimiter, quotechar='"')
        total_points_counter = 0
        for row in spamreader:
            #     0    1     2       3          4        5          6
            #   TYPE,SCOPE,NAME,DESCRIPTION,DATATYPE,SPECIFIER,ATTRIBUTES
            #   TYPE?SCOPE?NAME?DESCRIPTION?DATATYPE?SPECIFIER
            try:
                TYPE, SCOPE, NAME, DESCRIPTION, DATATYPE, SPECIFIER = row[0], row[1], row[2], row[3], row[
                    4], row[5],
            except IndexError:
                continue  # short string
            if TYPE == 'ALIAS':
                mapped_specifier = map_func(SPECIFIER)
                if not mapped_specifier == SPECIFIER:
                    # NAME = bold(NAME)
                    pass
                io_address = mapped_specifier.split(':', 2)
                if io_address[0] == 'RIO_SD' and False:
                    print(SPECIFIER, NAME)
                if len(io_address) == 3:  # ['RIO2_B', '8', 'I.Ch1Data'] or ['RIO_SD', '0', 'I.4']
                    chass, slot = io_address[0], int(io_address[1])
                    append_chass(chass_name=chass, slot_num=int(slot))

                    last_part = io_address[2].split('.', 2)
                    # I.Ch1Data I.Ch10Data vs I.Data.0 I.Data.10
                    if len(last_part) == 3:  # 'I.Data.10'
                        if last_part[0] == 'C':  # C.Ch0Config.HighEngineering
                            continue
                        if last_part[0] in 'IO' and last_part[1] == 'Data':
                            point = int(last_part[2])
                            # print(f'{tag_name} = {chass} {slot} {point}')
                            io_config[chass][slot][point] = NAME
                            io_description[chass][slot][point] = RUS_comment_decoder(DESCRIPTION)
                            total_points_counter += 1

                    if len(last_part) == 2:  # 'I.Ch3Data'  or   'I.4' for FlexIO
                        if last_part[0] == 'I' or last_part[0] == 'O':
                            if last_part[1].endswith('Data'):  # Ch8Data
                                point = last_part[1].removesuffix('Data').removeprefix('Ch')
                                if point.isdigit():
                                    point = int(point)
                                    # print(f'{tag_name} = {chass} {slot} {point}')
                                    # ic(io_config)
                                    io_config[chass][slot][point] = NAME
                                    io_description[chass][slot][point] = RUS_comment_decoder(DESCRIPTION)
                                    total_points_counter += 1
                                else:
                                    print(f'Unknown IO point format {last_part[1]}')
                            elif last_part[1].isdigit():  # just '4'  for FlexIO
                                point = int(last_part[1])
                                io_config[chass][slot][point] = NAME
                                io_description[chass][slot][point] = RUS_comment_decoder(DESCRIPTION)
                                total_points_counter += 1

        print(f'Total: {total_points_counter} points found')


def read_input_l5x(l5x_path, map_file_name=None, test_run=False, debug=False):
    """Read tags from L5X XML file and fill io_config/io_description."""
    global io_config
    global io_description

    print(f"Reading L5X XML file: {l5x_path}")
    project = l5x.Project(l5x_path)
    print(f"L5X project loaded: {project}")

    if map_file_name:
        n11 = n11mapping(map_file_name)
        map_func = n11.replace
    else:
        map_func = lambda s: s

    total_points_counter = 0
    parsed_counter = 0
    skipped_counter = 0
    map_counter = 0

    # --- 1. Теги в программах ---
    for prog in project.programs.names:
        for tag_name in project.programs[prog].tags.names:
            try:
                tag = project.programs[prog].tags[tag_name]
                alias_source = getattr(tag, 'alias_for', None)
                if alias_source:
                    alias = map_func(alias_source)
                    if alias != alias_source:
                        map_counter += 1
                else:
                    continue

                description = RUS_comment_decoder(getattr(tag, 'description', ""))
            except RuntimeError:
                continue

            if alias and ':' in alias:
                ok = process_alias_tag(f"{prog}/{tag_name}", alias, description, map_func, debug)
                if ok:
                    parsed_counter += 1
                else:
                    skipped_counter += 1
                total_points_counter += 1

    # --- 2. Контроллерные теги ---
    for tag_name in project.controller.tags.names:
        # trap for debug
        if tag_name == 'XA_DC1':
            pass
        try:
            tag = project.controller.tags[tag_name]
            alias_source = getattr(tag, 'alias_for', None)
            if alias_source:
                alias = map_func(alias_source)
                if alias != alias_source:
                    map_counter += 1
            else:
                continue
            description = RUS_comment_decoder(getattr(tag, 'description', ""))
        except RuntimeError:
            continue

        if alias and ':' in alias:
            ok = process_alias_tag(tag_name, alias, description, map_func, debug)
            if ok:
                parsed_counter += 1
            else:
                skipped_counter += 1
            total_points_counter += 1

    print(f"\nTotal {total_points_counter} alias tags processed.")
    print(f"  ✅ Parsed successfully: {parsed_counter}")
    print(f"  ⚠️  Skipped (unrecognized format): {skipped_counter}")
    print(f"  {map_counter} tags was mapped")


def process_alias_tag(tag_name, alias, description, map_func, debug=False):
    """Parse IO alias address (supports RIO, FlexBus, and short formats)."""
    global io_config, io_description

    alias_mapped = map_func(alias)
    parts = alias_mapped.split(':')

    chass = None
    slot = None
    path = None

    # --- Вариант 1: стандартный RIO_xx:x:O.Data.0 ---
    if len(parts) == 3:
        chass, slot_str, path = parts
        try:
            slot = int(slot_str)
            append_chass(chass, slot)
        except ValueError:
            if debug:
                print(f"  ❌ Skipped [{tag_name}] — invalid slot number: {alias_mapped}")
            return False

    # --- Вариант 2: короткий формат SD_Console:I.Data[0].0 ---
    elif len(parts) == 2:
        chass, path = parts
        slot = None  # определяем ниже из Data[...]
        if debug:
            print(f"  🟡 Detected short format [{alias_mapped}], slot будет определён из [{path}]")

    else:
        if debug:
            print(f"  ❌ Skipped [{tag_name}] — invalid alias format: {alias_mapped}")
        return False

    # --- Распознавание каналов и слотов (включая FlexBus) ---
    # Поддерживаем:
    #   I.0 / O.15
    #   I.Data.3 / O.Data.15
    #   I.Ch14Data / O.Ch14Data
    #   I.Ch[2].Data / O.Ch[2].Data
    #   O.Data[1].0 / I.Data[3].15   ← FlexBus: [1] — слот, .0 — канал

    flex_slot = None
    point = None

    # шаблон для разных форматов
    match = re.search(r"""
        ^[IO]\.?(
            (?P<num1>\d{1,3})$                            | # I.0
            [Dd]ata\.(?P<num2>\d{1,3})$                   | # I.Data.3
            [Dd]ata\[(?P<flex>\d{1,3})\]\.(?P<num3>\d{1,3})$ | # O.Data[1].0  ← FlexBus
            (?:Ch(?:annel)?\[?(?P<num4>\d{1,3})\]?(?:Data|\.[Dd]ata)?)$  # I.Ch14Data / I.Ch[2].Data
        )
    """, path, re.IGNORECASE | re.VERBOSE)

    if not match:
        if debug:
            print(f"  ⚠️  Skipped non-IO tag [{tag_name}] → {alias_mapped}")
        return False

    if match.group("flex"):
        flex_slot = int(match.group("flex"))
        point = int(match.group("num3"))
        append_chass(chass, flex_slot)

    else:
        point = int(
            match.group("num1") or
            match.group("num2") or
            match.group("num4")
        )

    # 🚫 исключаем служебные поля
    if re.search(r"(Fault|Status|Cfg|Config)", path, re.IGNORECASE):
        if debug:
            print(f"  🚫 Skipped service tag [{tag_name}] → {alias_mapped}")
        return False

    # --- запоминаем ---
    if flex_slot is not None:
        key = flex_slot
    else:
        key = slot

    try:
        if io_config[chass][key][point]:
            print(f"   Tag [{io_config[chass][key][point]}] replaced by [{tag_name}]")
    except KeyError:
        pass

    io_config[chass][key][point] = tag_name
    io_description[chass][key][point] = description

    if debug:
        fs = f" FlexSlot={flex_slot}" if flex_slot is not None else ""
        print(f"  ✅ Parsed [{tag_name}] → {chass}:{slot}:{point}{fs} ({path})")

    return True


def write_table(print_to_stdout=True):
    global io_config
    ms = f"""Created {datetime.datetime.now().isoformat()}
"""
    project_chass = list(io_config.keys())
    project_chass.sort()
    # ic(project_chass)
    for CHASSI in project_chass:
        cn = f'CHASSIS {CHASSI}'
        ms += f"""

{cn: ^125} 
╒══╤═════════════════╤═════════════════╤═════════════════╤═════════════════╤═════════════════╤═════════════════╤═════════════════╤═════════════════╤═════════════════╤═════════════════╕
│ch│      SLOT 0     │      SLOT 1     │      SLOT 2     │      SLOT 3     │      SLOT 4     │      SLOT 5     │      SLOT 6     │      SLOT 7     │      SLOT 8     │      SLOT 9     │ 
├──┼─────────────────┼─────────────────┼─────────────────┼─────────────────┼─────────────────┼─────────────────┼─────────────────┼─────────────────┼─────────────────┼─────────────────┤"""
        for CHANNEL in range(16):
            ms += f"""
│{CHANNEL:02}│"""
            for SLOT in range(0, 10):
                try:
                    tag = io_config[CHASSI][SLOT][CHANNEL]
                    tag = tag2kip(tag)
                except KeyError:
                    tag = ''
                ms += f"{tag: >17}│"
        ms += '''
└──┴─────────────────┴─────────────────┴─────────────────┴─────────────────┴─────────────────┴─────────────────┴─────────────────┴─────────────────┴─────────────────┴─────────────────┘
'''
    if print_to_stdout:
        print(ms)
    return ms


def write_table_compact():
    global io_config
    global io_description
    project_chass = list(io_config.keys())
    project_chass.sort()
    # ic(project_chass)
    ms = f"""Created {datetime.datetime.now().isoformat()}
"""
    for CHASSI in project_chass:
        cn = f'CHASSIS {CHASSI}'
        ms += f"""

{cn:=^22}
"""
        slot_list = list(io_config[CHASSI].keys())
        slot_list.sort()
        for SLOT in io_config[CHASSI].keys():
            ms += f"""
╒══╤═════════════════╕
│ch│     SLOT {SLOT:02}     │
├──┼─────────────────┤"""
            for CHANNEL in io_config[CHASSI][SLOT].keys():
                tag = io_config[CHASSI][SLOT][CHANNEL]
                descr = io_description.get(CHASSI, {}).get(SLOT, {}).get(CHANNEL, '')
                ms += f"""
│{CHANNEL:02}│{tag: >17}│ {descr}"""
            ms += f'''
└──┴─────────────────┘'''
    print(ms)


def write_csv_cspt(sep=','):
    """
    write datas in csv format
    Chassis, Slot, Point, Tagname
    :return:
    """
    global io_config
    project_chass = list(io_config.keys())
    project_chass.sort()
    # ic(project_chass)
    ms = f"""Created {datetime.datetime.now().isoformat()}
Chassis{sep}Slot{sep}Point,Tagname
"""
    for CHASSI in project_chass:
        slot_list = list(io_config[CHASSI].keys())
        slot_list.sort()
        for SLOT in io_config[CHASSI].keys():
            for CHANNEL in io_config[CHASSI][SLOT].keys():
                tag = io_config[CHASSI][SLOT][CHANNEL]
                ms += f"""
{CHASSI}{sep}{SLOT}{sep}{CHANNEL},{tag2kip(io_config[CHASSI][SLOT][CHANNEL])}"""
    print(ms)


def write_xlsx(out_file_name):
    global io_config
    print(f'xlsx writer selected. filename = {out_file_name}')
    workbook = xlsxwriter.Workbook(out_file_name)
    if True:
        # Add a formats.
        bold = workbook.add_format({'bold': True})
        slot_number_format = workbook.add_format()
        slot_number_format.set_bold()
        slot_number_format.set_font_color('gray')
        slot_number_format.set_align('center')
        slot_number_format.set_top(1)
        slot_number_format.set_bottom(1)

        ch_number_format = workbook.add_format()
        ch_number_format.set_align('center')
        ch_number_format.set_left(1)
        ch_number_format.set_right(1)

        right_to_content = workbook.add_format()
        right_to_content.set_right(1)

        content_format = workbook.add_format()
        content_format.set_center_across()

        date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})

    worksheet = workbook.add_worksheet()

    worksheet.write_string(0, 0, 'Created at')
    worksheet.write_datetime(0, 1, datetime.datetime.now(), date_format)

    worksheet.write_string(1, 0, 'Original input file name')
    worksheet.write_string(1, 1, str(out_file_name))

    # ==================================================================================================================
    row = 3

    project_chass = list(io_config.keys())
    project_chass.sort()

    def col_number(slot_number):
        return slot_number * 4 + 3

    def write_slot(_col, _row, slot_num, slot_data, descr_data={}):
        worksheet.write_string(_row, _col + 1, f'SLOT', slot_number_format)
        worksheet.write_number(_row, _col + 2, slot_num, slot_number_format)
        worksheet.write_blank(_row, _col + 3, '', slot_number_format)

        worksheet.write_blank(_row + 1, _col + 1, f'SLOT', slot_number_format)
        worksheet.write_blank(_row + 1, _col + 2, slot_num, slot_number_format)
        worksheet.write_blank(_row + 1, _col + 3, '', slot_number_format)

        worksheet.write_blank(_row, _col, '', ch_number_format)
        worksheet.write_blank(_row + 1, _col, '', ch_number_format)

        if len(slot_data.keys()) == 0:
            max_channel = 15
        else:
            max_channel = max(slot_data.keys())

        if max_channel <= 15:
            max_channel = 15
        else:
            max_channel = 31

        for Y in range(max_channel + 1):

            worksheet.write_number(_row + Y + 2, _col, Y, ch_number_format)
            tag = slot_data.get(Y, '')
            descr = descr_data.get(Y, '')
            worksheet.write_string(_row + Y + 2, _col + 1, tag2kip(tag), content_format)
            if descr:
                worksheet.write_comment(_row + Y + 2, _col + 1, descr.replace('$N', '\r'))

        worksheet.set_column(_col, _col, width=2.30)
        worksheet.set_column(_col + 1, _col + 1, width=23)

        return max_channel

    # ic(project_chass)
    for CHASSI in project_chass:
        row += 2
        worksheet.write_string(row, 0, f'CHASSIS')
        worksheet.write_string(row, 1, CHASSI, bold)
        row += 1
        size = 0
        for slot_num in range(13):  # slot numbers
            size = max(
                size,
                write_slot(col_number(slot_num),
                           row,
                           slot_num,
                           io_config[CHASSI].get(slot_num, {}),
                           io_description[CHASSI].get(slot_num, {})
                           )
            )
            # worksheet.write_number(row, col_number(slot_num+1), slot_num, slot_number_format)
            # worksheet.write_blank(row, col_number(slot_num), '', slot_number_format)  # left to slot number
            # worksheet.set_column(col_number(slot_num), col_number(slot_num), width=2.30)
            # worksheet.write_blank(row, col_number(slot_num) + 1, '', slot_number_format)  # right to slot number
            # worksheet.write_blank(row, col_number(slot_num) + 2, '', slot_number_format)  # right to slot number
            # worksheet.write_blank(row, col_number(slot_num) + 3, '', slot_number_format)  # right to slot number

        row += size + 2
    #     for CHANNEL in range(16):
    #         worksheet.write_number(row, 1, CHANNEL, ch_number_format)  # channel number
    #         for SLOT in range(0, 10):
    #             col = col_number(SLOT)
    #             try:
    #                 tag = io_config[CHASSI][SLOT][CHANNEL]
    #                 tag = tag2kip(tag)
    #             except KeyError:
    #                 tag = ''
    #             worksheet.write_string(row, col, tag, content_format)
    #             worksheet.write_blank(row, col + 1, '', right_to_content)
    #
    #         row += 1
    # # ==================================================================================================================

    workbook.read_only_recommended()
    workbook.close()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert CSV Controller Tags into a human readable table'
    )
    parser.add_argument('input_file', nargs='?', help="CSV or L5X file exported from RSLogix / Studio 5000")
    # parser.add_argument('input_csv', nargs='?', help="CSV file, exported from RSLogix")
    parser.add_argument('map', nargs='?', help="Substitution file (for N11/N68 mapping)")
    parser.add_argument('--debug', action='store_true', help="Show detailed tag parsing log")
    parser.add_argument('--old', action='store_true', help="CSV was generated by old version of RSLogix")
    parser.add_argument('--noxls', action='store_true', help="Do not write XLSX file")
    parser.add_argument('--test_run', action='store_true', help="Run test")
    parser.add_argument('--print', action='store_true', help="Print table to stdout")
    parser.add_argument('--print_compact', action='store_true', help="Print compact table to stdout")
    parser.add_argument('--version-info', action='store_true',
                        help="Show versions of xlsxwriter and l5x libraries")

    args = parser.parse_args()

    # если пользователь вызвал --version-info, просто выводим версии и выходим
    if args.version_info:
        print("Library versions:")
        print(f"  xlsxwriter: {getattr(xlsxwriter, '__version__', 'unknown')}")
        print(f"  l5x:        {getattr(l5x, '__version__', 'unknown')}")
        raise SystemExit(0)

    if args.test_run:
        print("Running tests")

    # ---- Проверка наличия входного файла ----
    if not args.input_file:
        parser.error("the following arguments are required: input_file")

    input_path = Path(args.input_file)
    if not input_path.is_file():
        print('No input file!')
        raise SystemExit(1)

    if args.map is None:
        print('No mapping will be used')
    elif not Path(args.map).is_file():
        print('No mapping file!')
        raise SystemExit(1)
    else:
        pass
        # print(f'{args.map} mapping will be used')

    # ---- Обработка по типу файла ----
    ext = input_path.suffix.lower()
    if ext == '.csv':
        print("Detected CSV input file.")
        if args.map is None:
            print('No mapping will be used')
        elif not Path(args.map).is_file():
            print('No mapping file!')
            raise SystemExit(1)
        read_input_csv(args.input_file, args.map, old_csv_version=args.old)

    elif ext == '.l5x':
        print("Detected L5X input file.")
        read_input_l5x(args.input_file, map_file_name=args.map, test_run=args.test_run, debug=args.debug)

    else:
        print(f"Unsupported file type: {ext}")
        raise SystemExit(1)

    if args.print_compact:
        write_table_compact()
    if args.print:
        write_table()
    # write_table_compact()
    # write_csv_cspt(sep=':')
    if not args.noxls:
        out_xlsx = input_path.with_suffix('.xlsx')
        write_xlsx(out_xlsx)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
