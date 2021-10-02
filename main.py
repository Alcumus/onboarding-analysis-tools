import argparse
import csv
import re
from fuzzywuzzy import fuzz
CBX_ID = 0
CBX_COMPANY_FR = 1
CBX_COMPANY_EN = 2
CBX_COMPANY_OLD = 3
CBX_ADDRESS = 4
CBX_CITY = 5
CBX_STATE = 6
CBX_COUNTRY = 7
CBX_ZIP = 8
CBX_FISTNAME = 9
CBX_LASTNAME = 10
CBX_EMAIL = 11
CBX_EXPIRATION_DATE = 12
CBX_REGISTRATION_STATUS = 13
CBX_SUSPENDED = 14
CBX_MODULES = 15
CBX_ACCOUNT_TYPE = 16
CBX_SUB_PRICE = 17
CBX_EMPL_PRICE = 18
CBX_HIRING_CLIENT_NAMES = 19
CBX_HIRING_CLIENT_IDS = 20
CBX_HIRING_CLIENT_QSTATUS = 21
CBX_PARENTS = 22

HC_COMPANY = 0
HC_FIRSTNAME = 1
HC_LASTNAME = 2
HC_EMAIL = 3
HC_PHONE_NUMBER = 4
HC_LANGUAGE = 5
HC_STREET = 6
HC_CITY = 7
HC_STATE = 8
HC_COUNTRY = 9
HC_ZIP = 10
HC_CATEGORY = 11
HC_IS_TAKE_OVER = 12
HC_TAKEOVER_RENEWAL_DATE = 13
HC_TAKEOVER_QF_STATUS = 14
HC_PROJECT_NAME = 15
HC_QUESTIONNAIRE_NAME = 16
HC_QUESTIONNAIRE_ID = 17
HC_CONTRACTOR_ACCOUNT_TYPE = 18
HC_HIRING_CLIENT_NAME = 19
HC_HIRING_CLIENT_ID = 20
HC_IS_ASSOCIATION_FEE = 21
HC_BASE_SUBSCRIPTION_FEE = 22
HC_DO_NOT_MATCH = 23
HC_FORCE_CBX_ID = 24
HC_AMBIGUOUS = 25


#todo fix the subcription price
#todo classify business units



BASE_GENERIC_DOMAIN = ['yahoo.ca', 'yahoo.com', 'hotmail.com', 'gmail.com', 'outlook.com',
                       'bell.com', 'bell.ca', 'videotron.ca', 'eastlink.ca', 'kos.net', 'bellnet.ca', 'sasktel.net',
                       'aol.com', 'tlb.sympatico.ca', 'sogetel.net', 'cgocable.ca',
                       'hotmail.ca', 'live.ca', 'icloud.com', 'hotmail.fr', 'yahoo.com', 'outlook.fr', 'msn.com',
                       'globetrotter.net', 'live.com', 'sympatico.ca', 'live.fr', 'yahoo.fr', 'telus.net',
                       'shaw.ca', 'me.com', 'bell.net',
                       '']

# define commandline parser
parser = argparse.ArgumentParser(description='Tool to match contractor list provided by hiring clients to contractors in CBX, '
                                             'all input/output files must be in the current directory',
                                 formatter_class=argparse.RawTextHelpFormatter)
parser.add_argument('cbx_list',
                    help='''csv DB export file (no header) of contractors with the following columns: 
    TBD''')

parser.add_argument('hc_list',
                    help='''csv file (with header) and the following columns:
    contractor, firstname, lastname, email, add_email1, add_email2, street, city, state, country, zip code, category, is take over, take over renewal date'''
                    )
parser.add_argument('output',
                    help='''csv file with the following columns: 
    <<hc_list columns>>, Cognibox ID, contractor, company name score, contact score, address score,
    is CBX, is active member, is take over, is subscription upgrade, is association fee 
    matching information  
Matching information format:
    Cognibox ID, firstname lastname, birthdate --> Contractor 1 [parents: C1 parent1;C1 parent2;etc..] 
    [previous: Empl. Previous1;Empl. Previous2], match ratio 1,
    Contractor 2 [C2 parent1;C2 parent2;etc..], match ratio 2, etc...
The matching ratio is a value betwween 0 and 100, where 100 is a perfect match.
Please note the Cognibox ID and birthdate is set ONLY if a single match his found. If no match
or multiple matches are found it is left empty.''')

parser.add_argument('--cbx_list_encoding', dest='cbx_encoding', action='store',
                    default='utf-8-sig',
                    help='Encoding for the cbx list (default: utf-8-sig)')

parser.add_argument('--hc_list_encoding', dest='hc_encoding', action='store',
                    default='utf-8-sig',
                    help='Encoding for the hc list (default: utf-8-sig)')

parser.add_argument('--output_encoding', dest='output_encoding', action='store',
                    default='utf-8-sig',
                    help='Encoding for the hc list (default: utf-8-sig)')


parser.add_argument('--min_company_match_ratio', dest='ratio_company', action='store',
                    default=80,
                    help='Minimum match ratio for contractors, between 0 and 100 (default 60)')

parser.add_argument('--list_separator', dest='list_separator', action='store',
                    default=';',
                    help='string separator used for lists (default: ;)')

parser.add_argument('--additional_generic_domain', dest='additional_generic_domain', action='store',
                    default='',
                    help='list of domains to ignore separated by the list separator (default separator is ;)')

parser.add_argument('--min_address_match_ratio', dest='ratio_address', action='store',
                    default=80,
                    help='Minimum match ratio for contractors, between 0 and 100 (default 70)')

args = parser.parse_args()

GENERIC_DOMAIN = BASE_GENERIC_DOMAIN + args.additional_generic_domain.split(args.list_separator)


def add_analysis_data(hc_row, cbx_row, ratio_company=None, ratio_address=None, contact_match=None):
    cbx_company = cbx_row[CBX_COMPANY_FR] if cbx_row[CBX_COMPANY_FR] else cbx_row[CBX_COMPANY_EN]
    print('   --> ', cbx_company, hc_email, cbx_row[CBX_ID], ratio_company, ratio_address, contact_match)
    hiring_clients_list = cbx_row[CBX_HIRING_CLIENT_NAMES].split(args.list_separator)
    hiring_clients_qstatus = cbx_row[CBX_HIRING_CLIENT_QSTATUS].split(args.list_separator)
    hc_count = len(hiring_clients_list) if cbx_row[CBX_HIRING_CLIENT_NAMES] else 0
    is_in_relationship = True if (
                hc_row[HC_HIRING_CLIENT_NAME] in hiring_clients_list and hc_row[HC_HIRING_CLIENT_NAME]) else False
    hc_index = 0
    is_qualified = False
    while hc_index < len(hiring_clients_list):
        if hiring_clients_list[hc_index] == hc_row[HC_HIRING_CLIENT_NAME] and \
                hiring_clients_qstatus[hc_index] == 'validated':
            is_qualified = True
            break
        hc_index += 1
    return {'cbx_id': cbx_row[CBX_ID],
            'company': cbx_company,
            'address': cbx_row[CBX_ADDRESS],
            'city': cbx_row[CBX_CITY],
            'state': cbx_row[CBX_STATE],
            'zip': cbx_row[CBX_ZIP],
            'country': cbx_row[CBX_COUNTRY],
            'expiration_date': cbx_row[CBX_EXPIRATION_DATE],
            'registration_status': cbx_row[CBX_REGISTRATION_STATUS],
            'suspended': cbx_row[CBX_SUSPENDED],
            'email': cbx_row[CBX_EMAIL],
            'first_name': cbx_row[CBX_FISTNAME],
            'last_name': cbx_row[CBX_LASTNAME],
            'modules': cbx_row[CBX_MODULES],
            'account_type': cbx_row[CBX_ACCOUNT_TYPE],
            'subscription_price': cbx_row[CBX_SUB_PRICE],
            'employee_price': cbx_row[CBX_EMPL_PRICE],
            'is_subscription_upgrade': str(False),
            'parents': cbx_row[CBX_PARENTS],
            'previous': cbx_row[CBX_COMPANY_OLD],
            'hiring_client_names': cbx_row[CBX_HIRING_CLIENT_NAMES],
            'hiring_client_count': hc_count,
            'is_in_relationship': is_in_relationship,
            'is_qualified': str(is_qualified),
            'ratio_company': ratio_company,
            'ratio_address': ratio_address,
            'contact_match': str(contact_match),
            }

if __name__ == '__main__':
    data_path = './data/'
    cbx_file = data_path + args.cbx_list
    hc_file = data_path + args.hc_list
    output_file = data_path + args.output

    # output parameters used
    print(f'Reading CBX list: {args.cbx_list} [{args.cbx_encoding}]')
    print(f'Reading HC list: {args.hc_list} [{args.hc_encoding}]')
    print(f'Outputing results in: {args.output} [{args.output_encoding}]')
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
    headers = cbx_data.pop(0)
    print(f'Completed reading {len(cbx_data)} contractors.')
    print('Reading hiring client data file...')
    with open(hc_file, 'r', encoding=args.hc_encoding) as hc:
        for row in csv.reader(hc):
            hc_data.append(row)
    print(f'Completed reading {len(hc_data)} contractors.')
    total = len(hc_data) - 1
    index = 1
    with open(output_file, 'w', newline='', encoding=args.output_encoding) as resultfile:
        writer = csv.writer(resultfile)
        headers = hc_data.pop(0)
        metadata_indexes = []
        h_index = 0
        while(h_index < len(headers)):
            if headers[h_index].lower().startswith('metadata'):
                metadata_indexes.append(h_index)
            h_index += 1
        metadata_indexes.sort(reverse=True)
        #todo check headers names
        headers.extend(['cbx_id', 'cbx_contractor', 'cbx_street', 'cbx_city', 'cbx_state', 'cbx_zip', 'cbx_country',
                        'expiration_date', 'registration_status', 'suspended', 'cbx_email',
                        'cbx_first_name', 'cbx_last_name', 'modules', 'cbx_account_type',
                        'cbx_subscription_fee', 'cbx_employee_price',  'is_subscription_upgrade', 'parents', 'previous',
                        'hiring_client_names', 'hiring_client_count',
                        'is_in_relationship', 'is_qualified', 'ratio_company', 'ratio_address',
                        'contact_match', 'generic_domain', 'match_count', 'match_count_with_hc',
                        'analysis', 'index'])
        metadata_array = []
        for md_index in metadata_indexes:
            metadata_array.insert(0, headers.pop(md_index))
        headers.extend(metadata_array)
        writer.writerow(headers)
        # match
        for hc_row in hc_data:
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
                        matches.append(add_analysis_data(hc_row,cbx_row))
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
                        ratio_company_fr = fuzz.token_sort_ratio(cbx_company_fr.lower().replace('.', '').replace(',', '').strip(),
                                                                 clean_hc_company)
                        ratio_company_en = fuzz.token_sort_ratio(cbx_company_en.lower().replace('.', '').replace(',', '').strip(),
                                                                 clean_hc_company)
                        ratio_address = fuzz.token_sort_ratio(cbx_address,
                                                              hc_address)
                        ratio_address = ratio_address if ratio_zip == 0 else ratio_zip if ratio_address == 0 else ratio_address * ratio_zip / 100
                        ratio_company = ratio_company_fr if ratio_company_fr > ratio_company_en else ratio_company_en
                        ratio_previous = 0
                        for item in cbx_previous.split(args.list_separator):
                            if item in (cbx_company_en, cbx_company_fr):
                                continue
                            ratio = fuzz.token_sort_ratio(item.lower().replace('.', '').replace(',', '').strip(),
                                                          clean_hc_company)
                            ratio_previous = ratio if ratio > ratio_previous else ratio_previous
                        ratio_company = ratio_previous if ratio_previous > ratio_company else ratio_company
                        if ((ratio_company >= float(args.ratio_company) and ratio_address >= float(args.ratio_address)) or
                                contact_match):
                            matches.append(add_analysis_data(hc_row, cbx_row, ratio_company, ratio_address, contact_match))
            ids = []
            best_match = 0
            matches = sorted(matches, key=lambda x: (x['hiring_client_count'], x["ratio_address"], x["ratio_company"]), reverse=True)
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
            hc_row.append(str(index))
            metadata_array = []
            for md_index in metadata_indexes:
                metadata_array.insert(0,hc_row.pop(md_index))
            hc_row.extend(metadata_array)
            writer.writerow(hc_row)
            print(f'{index} of {total} [{len(uniques_cbx_id)} found]')
            index += 1
