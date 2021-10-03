import argparse
import csv
import re
from fuzzywuzzy import fuzz

# noinspection SpellCheckingInspection
CBX_ID, CBX_COMPANY_FR, CBX_COMPANY_EN, CBX_COMPANY_OLD, CBX_ADDRESS, CBX_CITY, CBX_STATE, \
    CBX_COUNTRY, CBX_ZIP, CBX_FISTNAME, CBX_LASTNAME, CBX_EMAIL, CBX_EXPIRATION_DATE, CBX_REGISTRATION_STATUS, \
    CBX_SUSPENDED, CBX_MODULES, CBX_ACCOUNT_TYPE, CBX_SUB_PRICE, CBX_EMPL_PRICE, CBX_HIRING_CLIENT_NAMES, \
    CBX_HIRING_CLIENT_IDS, CBX_HIRING_CLIENT_QSTATUS, CBX_PARENTS = range(23)

HC_COMPANY, HC_FIRSTNAME, HC_LASTNAME, HC_EMAIL, HC_PHONE_NUMBER, HC_LANGUAGE, HC_STREET, HC_CITY, \
    HC_STATE, HC_COUNTRY, HC_ZIP, HC_CATEGORY, HC_IS_TAKE_OVER, HC_TAKEOVER_RENEWAL_DATE, HC_TAKEOVER_QF_STATUS, \
    HC_PROJECT_NAME, HC_QUESTIONNAIRE_NAME, HC_QUESTIONNAIRE_ID, HC_CONTRACTOR_ACCOUNT_TYPE, HC_HIRING_CLIENT_NAME, \
    HC_HIRING_CLIENT_ID, HC_IS_ASSOCIATION_FEE, HC_BASE_SUBSCRIPTION_FEE, HC_DO_NOT_MATCH, HC_FORCE_CBX_ID, \
    HC_AMBIGUOUS = range(26)

# noinspection SpellCheckingInspection
cbx_headers = ['id', 'name_fr', 'name_en', 'old_names', 'address', 'city', 'state', 'country', 'postal_code',
               'first_name', 'last_name', 'email', 'cbx_expiration_date', 'registration_code', 'suspended',
               'modules', 'code', 'subscription_price', 'employee_price', 'hiring_client_names',
               'hiring_client_ids', 'hiring_client_qstatus', 'parents']

# noinspection SpellCheckingInspection
hc_headers = ['contractor', 'first name', 'last name', 'email', 'phone_number', 'language', 'street', 'city',
              'state', 'country', 'zip', 'category', 'is_take_over', 'take_over_renewal_date',
              'take_over_qualification_status', 'project_name', 'questionnaire_name', 'questionnaire_id',
              'account_type', 'hiring_client_name', 'hiring_client_id', 'is_association_fee',
              'base_subscription_fee', 'do_not_match', 'force_cbx_id', 'ambiguous']

# noinspection SpellCheckingInspection
analysis_headers = ['cbx_id', 'cbx_contractor', 'cbx_street', 'cbx_city', 'cbx_state', 'cbx_zip', 'cbx_country',
                    'expiration_date', 'registration_status', 'suspended', 'cbx_email',
                    'cbx_first_name', 'cbx_last_name', 'modules', 'cbx_account_type',
                    'cbx_subscription_fee', 'cbx_employee_price', 'is_subscription_upgrade', 'parents', 'previous',
                    'hiring_client_names', 'hiring_client_count',
                    'is_in_relationship', 'is_qualified', 'ratio_company', 'ratio_address',
                    'contact_match', 'generic_domain', 'match_count', 'match_count_with_hc',
                    'analysis', 'index']

metadata_headers = ['metadata_x', 'metadata_y', 'metadata_z', '...']

# todo fix the subscription price
# todo classify business units

# noinspection SpellCheckingInspection
BASE_GENERIC_DOMAIN = ['yahoo.ca', 'yahoo.com', 'hotmail.com', 'gmail.com', 'outlook.com',
                       'bell.com', 'bell.ca', 'videotron.ca', 'eastlink.ca', 'kos.net', 'bellnet.ca', 'sasktel.net',
                       'aol.com', 'tlb.sympatico.ca', 'sogetel.net', 'cgocable.ca',
                       'hotmail.ca', 'live.ca', 'icloud.com', 'hotmail.fr', 'yahoo.com', 'outlook.fr', 'msn.com',
                       'globetrotter.net', 'live.com', 'sympatico.ca', 'live.fr', 'yahoo.fr', 'telus.net',
                       'shaw.ca', 'me.com', 'bell.net',
                       '']


def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


hc_headers_with_metadata = hc_headers.copy()
hc_headers_with_metadata.extend(metadata_headers)
cbx_headers_text = '\n'.join([', '.join(x) for x in list(chunks(cbx_headers, 5))])
hc_headers_text = '\n'.join([', '.join(x) for x in list(chunks(hc_headers_with_metadata, 5))])
analysis_headers_text = '\n'.join([', '.join(x) for x in list(chunks(analysis_headers, 5))])

# define commandline parser
parser = argparse.ArgumentParser(
    description='Tool to match contractor list provided by hiring clients to business units in CBX, '
                'all input/output files must be in the current directory',
    formatter_class=argparse.RawTextHelpFormatter)
parser.add_argument('cbx_list',
                    help=f'csv DB export file of business units with the following columns:\n{cbx_headers_text}\n\n')

parser.add_argument('hc_list',
                    help=f'csv file of the hiring client contractors and the following columns:\n{hc_headers_text}\n\n')
parser.add_argument('output',
                    help=f'csv file with the hc_list columns and the following analysis columns:'
                         f'\n{analysis_headers_text}\n\n**Please note that metadata columns from the'
                         f' hc file are moved after the analysis data')

parser.add_argument('--min_company_match_ratio', dest='ratio_company', action='store',
                    default=80,
                    help='Minimum match ratio for contractors, between 0 and 100 (default 80)')

parser.add_argument('--min_address_match_ratio', dest='ratio_address', action='store',
                    default=80,
                    help='Minimum match ratio for addresses (street + zip), between 0 and 100 (default 80)')

parser.add_argument('--additional_generic_domain', dest='additional_generic_domain', action='store',
                    default='',
                    help='list of domains to ignore separated by the list separator (default separator is ;)')

parser.add_argument('--cbx_list_encoding', dest='cbx_encoding', action='store',
                    default='utf-8-sig',
                    help='Encoding for the cbx list (default: utf-8-sig)')

parser.add_argument('--hc_list_encoding', dest='hc_encoding', action='store',
                    default='utf-8-sig',
                    help='Encoding for the hc list (default: utf-8-sig)')

parser.add_argument('--output_encoding', dest='output_encoding', action='store',
                    default='utf-8-sig',
                    help='Encoding for the hc list (default: utf-8-sig)')

parser.add_argument('--list_separator', dest='list_separator', action='store',
                    default=';',
                    help='string separator used for lists (default: ;)')

parser.add_argument('--no_headers', dest='no_headers', action='store_true',
                    help='to indicate that input files have no headers')

parser.add_argument('--ignore_warnings', dest='ignore_warnings', action='store_true',
                    help='to ignore data consistency checks and run anyway...')

args = parser.parse_args()
GENERIC_DOMAIN = BASE_GENERIC_DOMAIN + args.additional_generic_domain.split(args.list_separator)


# noinspection PyShadowingNames
def add_analysis_data(hc_row, cbx_row, ratio_company=None, ratio_address=None, contact_match=None):
    cbx_company = cbx_row[CBX_COMPANY_FR] if cbx_row[CBX_COMPANY_FR] else cbx_row[CBX_COMPANY_EN]
    print('   --> ', cbx_company, hc_email, cbx_row[CBX_ID], ratio_company, ratio_address, contact_match)
    hiring_clients_list = cbx_row[CBX_HIRING_CLIENT_NAMES].split(args.list_separator)
    hiring_clients_qstatus = cbx_row[CBX_HIRING_CLIENT_QSTATUS].split(args.list_separator)
    hc_count = len(hiring_clients_list) if cbx_row[CBX_HIRING_CLIENT_NAMES] else 0
    is_in_relationship = True if (
            hc_row[HC_HIRING_CLIENT_NAME] in hiring_clients_list and hc_row[HC_HIRING_CLIENT_NAME]) else False
    is_qualified = False
    for idx, val in enumerate(hiring_clients_list):
        if val == hc_row[HC_HIRING_CLIENT_NAME] and hiring_clients_qstatus[idx] == 'validated':
            is_qualified = True
            break
    return {'cbx_id': cbx_row[CBX_ID], 'company': cbx_company, 'address': cbx_row[CBX_ADDRESS],
            'city': cbx_row[CBX_CITY], 'state': cbx_row[CBX_STATE], 'zip': cbx_row[CBX_ZIP],
            'country': cbx_row[CBX_COUNTRY], 'expiration_date': cbx_row[CBX_EXPIRATION_DATE],
            'registration_status': cbx_row[CBX_REGISTRATION_STATUS],
            'suspended': cbx_row[CBX_SUSPENDED], 'email': cbx_row[CBX_EMAIL], 'first_name': cbx_row[CBX_FISTNAME],
            'last_name': cbx_row[CBX_LASTNAME], 'modules': cbx_row[CBX_MODULES],
            'account_type': cbx_row[CBX_ACCOUNT_TYPE], 'subscription_price': cbx_row[CBX_SUB_PRICE],
            'employee_price': cbx_row[CBX_EMPL_PRICE], 'is_subscription_upgrade': str(False),
            'parents': cbx_row[CBX_PARENTS], 'previous': cbx_row[CBX_COMPANY_OLD],
            'hiring_client_names': cbx_row[CBX_HIRING_CLIENT_NAMES], 'hiring_client_count': hc_count,
            'is_in_relationship': is_in_relationship, 'is_qualified': str(is_qualified),
            'ratio_company': ratio_company, 'ratio_address': ratio_address, 'contact_match': str(contact_match),
            }


# noinspection PyShadowingNames
def check_headers(headers, standards, ignore):
    headers = [x.lower().strip() for x in headers]
    for idx, val in enumerate(standards):
        if val != headers[idx]:
            print(f'WARNING: got "{headers[idx]}" while expecting "{val}" in column {idx + 1}')
            if not ignore:
                exit(-1)


if __name__ == '__main__':
    data_path = './data/'
    cbx_file = data_path + args.cbx_list
    hc_file = data_path + args.hc_list
    output_file = data_path + args.output

    # output parameters used
    print(f'Reading CBX list: {args.cbx_list} [{args.cbx_encoding}]')
    print(f'Reading HC list: {args.hc_list} [{args.hc_encoding}]')
    print(f'Outputting results in: {args.output} [{args.output_encoding}]')
    print(f'contractor match ratio: {args.ratio_company}')
    print(f'address match ratio: {args.ratio_address}')
    print(f'list of "generic domains:\n{BASE_GENERIC_DOMAIN}')
    print(f'additional generic domain: {args.additional_generic_domain}')
    # read data
    cbx_data = []
    hc_data = []
    print('Reading Cognibox data file...')
    with open(cbx_file, 'r', encoding=args.cbx_encoding) as cbx:
        for row in csv.reader(cbx):
            cbx_data.append(row)
    # check cbx db ata consistency
    if cbx_data and len(cbx_data[0]) != len(cbx_headers):
        print(f'WARNING: got {len(cbx_data[0])} columns when expecting {len(cbx_headers)}')
        if not args.ignore_warnings:
            exit(-1)
    if not args.no_headers:
        headers = cbx_data.pop(0)
        headers = [x.lower().strip() for x in headers]
        check_headers(headers, cbx_headers, args.ignore_warnings)
    print(f'Completed reading {len(cbx_data)} contractors.')

    print('Reading hiring client data file...')
    with open(hc_file, 'r', encoding=args.hc_encoding) as hc:
        for row in csv.reader(hc):
            hc_data.append(row)
    total = len(hc_data) - 1
    metadata_indexes = []
    headers = []
    # check hc data consistency
    if hc_data and len(hc_data[0]) < len(hc_headers):
        print(f'WARNING: got {len(hc_data[0])} columns when at least {len(hc_headers)} is expected')
        if not args.ignore_warnings:
            exit(-1)
    if not args.no_headers:
        headers = hc_data.pop(0)
        headers = [x.lower().strip() for x in headers]
        check_headers(headers, hc_headers, args.ignore_warnings)
    else:
        if hc_data and len(hc_data[0]) != len(hc_headers):
            print(f'WARNING: got {len(hc_data[0])} columns when {len(hc_headers)} is exactly expected')
            if not args.ignore_warnings:
                exit(-1)
    print(f'Completed reading {len(hc_data)} contractors.')

    with open(output_file, 'w', newline='', encoding=args.output_encoding) as result_file:
        writer = csv.writer(result_file)
        # append analysis headers and move metadata headers at the end
        if not args.no_headers:
            for idx, val in enumerate(headers):
                if val.lower().startswith('metadata'):
                    metadata_indexes.append(idx)
            metadata_indexes.sort(reverse=True)
            headers.extend(analysis_headers)
            metadata_array = []
            for md_index in metadata_indexes:
                metadata_array.insert(0, headers.pop(md_index))
            headers.extend(metadata_array)
            writer.writerow(headers)

        # match
        for index, hc_row in enumerate(hc_data):
            matches = []
            hc_company = hc_row[HC_COMPANY]
            clean_hc_company = hc_company.lower().replace('.', '').replace(',', '').strip()
            clean_hc_company = modified_string = re.sub(r"\([^()]*\)", "", clean_hc_company)
            hc_email = hc_row[HC_EMAIL].lower()
            hc_domain = hc_email[hc_email.find('@') + 1:]
            hc_zip = hc_row[HC_ZIP].replace(' ', '').upper()
            hc_address = hc_row[HC_STREET].lower().replace('.', '').strip()
            hc_force_cbx = hc_row[HC_FORCE_CBX_ID].strip()
            if hc_row[HC_DO_NOT_MATCH].lower().strip() != 'true' and hc_row[HC_DO_NOT_MATCH].strip() != '1':
                if hc_force_cbx:
                    cbx_row = next(filter(lambda x: x[CBX_ID].strip() == hc_force_cbx, cbx_data), None)
                    if cbx_row:
                        matches.append(add_analysis_data(hc_row, cbx_row))
                else:
                    for cbx_row in cbx_data:
                        cbx_email = cbx_row[CBX_EMAIL].lower()
                        cbx_domain = cbx_email[cbx_email.find('@') + 1:]
                        contact_match = False
                        if hc_domain in GENERIC_DOMAIN:
                            contact_match = True if cbx_email == hc_email else False
                        else:
                            contact_match = True if cbx_domain == hc_domain else False
                        cbx_zip = cbx_row[CBX_ZIP].replace(' ', '').upper()
                        cbx_company_en = re.sub(r"\([^()]*\)", "", cbx_row[CBX_COMPANY_EN])
                        cbx_company_fr = re.sub(r"\([^()]*\)", "", cbx_row[CBX_COMPANY_FR])
                        cbx_parents = cbx_row[CBX_PARENTS]
                        cbx_previous = cbx_row[CBX_COMPANY_OLD]
                        cbx_address = cbx_row[CBX_ADDRESS].lower().replace('.', '').strip()
                        ratio_zip = fuzz.ratio(cbx_zip, hc_zip)
                        ratio_company_fr = fuzz.token_sort_ratio(
                            cbx_company_fr.lower().replace('.', '').replace(',', '').strip(),
                            clean_hc_company)
                        ratio_company_en = fuzz.token_sort_ratio(
                            cbx_company_en.lower().replace('.', '').replace(',', '').strip(),
                            clean_hc_company)
                        ratio_address = fuzz.token_sort_ratio(cbx_address,
                                                              hc_address)
                        ratio_address = ratio_address if ratio_zip == 0 else ratio_zip if ratio_address == 0 \
                            else ratio_address * ratio_zip / 100
                        ratio_company = ratio_company_fr if ratio_company_fr > ratio_company_en else ratio_company_en
                        ratio_previous = 0
                        for item in cbx_previous.split(args.list_separator):
                            if item in (cbx_company_en, cbx_company_fr):
                                continue
                            ratio = fuzz.token_sort_ratio(item.lower().replace('.', '').replace(',', '').strip(),
                                                          clean_hc_company)
                            ratio_previous = ratio if ratio > ratio_previous else ratio_previous
                        ratio_company = ratio_previous if ratio_previous > ratio_company else ratio_company
                        if ((ratio_company >= float(args.ratio_company) and ratio_address >= float(
                                args.ratio_address)) or
                                contact_match):
                            matches.append(
                                add_analysis_data(hc_row, cbx_row, ratio_company, ratio_address, contact_match))
            ids = []
            best_match = 0
            matches = sorted(matches, key=lambda x: (x['hiring_client_count'], x["ratio_address"], x["ratio_company"]),
                             reverse=True)
            for item in matches[0:10]:
                ids.append(f'{item["cbx_id"]}, {item["company"]}, {item["address"]}, {item["zip"]}, {item["email"]}'
                           f' --> CR{item["ratio_company"]}, AR{item["ratio_address"]},'
                           f' CM{item["contact_match"]}, HCC{item["hiring_client_count"]}')
            # append matching results to the hc_list
            match_data = []
            uniques_cbx_id = set(item['cbx_id'] for item in matches)
            if uniques_cbx_id:
                for key, value in matches[0].items():
                    match_data.append(value)
                hc_row.extend(match_data)
                hc_row.append(str(True if hc_domain in GENERIC_DOMAIN else False))
                hc_row.append(len(uniques_cbx_id) if len(uniques_cbx_id) else '')
                hc_row.append(str(len([i for i in matches if i['hiring_client_count'] > 0])))
                hc_row.append('|'.join(ids))
            else:
                hc_row.extend(['' for x in range(31)])
            hc_row.append(str(index+1))
            metadata_array = []
            for md_index in metadata_indexes:
                metadata_array.insert(0, hc_row.pop(md_index))
            hc_row.extend(metadata_array)
            writer.writerow(hc_row)
            print(f'{index+1} of {total} [{len(uniques_cbx_id)} found]')
