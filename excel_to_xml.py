import pandas as pd
from lxml import etree
from datetime import datetime


def excel_to_xml(input_path, output_path):
    # Чтение данных компании
    company_df = pd.read_excel(input_path, sheet_name='Данные компании', header=None)
    company_data = {
        'org_name': company_df.iloc[1, 1],
        'LP_TIN': company_df.iloc[2, 1],
        'address': company_df.iloc[4, 1],
        'phone': company_df.iloc[5, 1],
        'country_code': company_df.iloc[6, 1],
        'units_per_pack': int(company_df.iloc[11, 1])  # Динамическое чтение из файла
    }

    # Чтение кодов наборов и единиц
    pack_codes = pd.read_excel(input_path, sheet_name='Коды наборов', header=None)[0].tolist()
    unit_codes = pd.read_excel(input_path, sheet_name='Коды единиц', header=None)[0].tolist()

    # Проверка соответствия количества
    if len(unit_codes) != len(pack_codes) * company_data['units_per_pack']:
        raise ValueError(f"Количество unit codes ({len(unit_codes)}) должно быть равно "
                         f"pack codes ({len(pack_codes)}) × {company_data['units_per_pack']}")

    # Создание XML структуры
    root = etree.Element('unit_pack', document_id="auto_generated", VerForm="1.03")
    doc = etree.SubElement(root, 'Document',
                           operation_date_time=datetime.now().isoformat(),
                           document_number="023")

    # Секция организации (остается без изменений)
    org = etree.SubElement(doc, 'organisation')
    id_info = etree.SubElement(org, 'id_info')
    etree.SubElement(id_info, 'LP_info',
                     org_name=company_data['org_name'],
                     LP_TIN=str(company_data['LP_TIN']),
                     RRC="nan")

    address = etree.SubElement(org, 'Address')
    etree.SubElement(address, 'location_address',
                     country_code=str(company_data['country_code']),
                     text_address=company_data['address'])

    contacts = etree.SubElement(org, 'contacts')
    etree.SubElement(contacts, 'phone_number').text = str(company_data['phone'])
    etree.SubElement(contacts, 'email').text = "example@example.com"

    # Генерация pack_content с динамическим количеством
    units_grouped = [unit_codes[i:i + company_data['units_per_pack']]
                     for i in range(0, len(unit_codes), company_data['units_per_pack'])]

    for pack_code, units in zip(pack_codes, units_grouped):
        pc = etree.SubElement(doc, 'pack_content')
        etree.SubElement(pc, 'pack_code').text = pack_code
        for unit in units:
            cis = etree.SubElement(pc, 'cis')
            cis.text = unit

    # Сохранение файла
    xml_str = etree.tostring(root, pretty_print=True, encoding='utf-8', xml_declaration=True)
    with open(output_path, 'wb') as f:
        f.write(xml_str)
