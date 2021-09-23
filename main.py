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
CBX_HIRING_CLIENTS = 19
CBX_PARENTS = 20

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
HC_QUESTIONNAIRE = 12
HC_CONTRACTOR_ACCOUNT_TYPE = 13
HC_HIRING_CLIENT_ID = 14
HC_IS_ASSOCIATION_FEE = 15
HC_BASE_SUBSCRIPTION_FEE = 16
HC_IS_TAKE_OVER = 17
HC_DO_NOT_MATCH = 18
HC_FORCE_CBX_ID = 19


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
    overall score,  is CBX, is active member, is take over, is subscription upgrade, is association fee 
    matching information  
Matching information format:
    Cognibox ID, firstname lastname, birthdate --> Contractor 1 [parents: C1 parent1;C1 parent2;etc..] 
    [previous: Empl. Previous1;Empl. Previous2], match ratio 1,
    Contractor 2 [C2 parent1;C2 parent2;etc..], match ratio 2, etc...
The matching ratio is a value betwween 0 and 100, where 100 is a perfect match.
Please note the Cognibox ID and birthdate is set ONLY if a single match his found. If no match
or multiple matches are found it is left empty.''')

parser.add_argument('--cbx_list_encoding', dest='cbx_encoding', action='store',
                    default='utf-8',
                    help='Encoding for the cbx list (default: utf-8)')

parser.add_argument('--hc_list_encoding', dest='hc_encoding', action='store',
                    default='cp1252',
                    help='Encoding for the hc list (default: cp1252)')

parser.add_argument('--output_encoding', dest='output_encoding', action='store',
                    default='cp1252',
                    help='Encoding for the hc list (default: cp1252)')


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
                    default=90,
                    help='Minimum match ratio for contractors, between 0 and 100 (default 70)')

args = parser.parse_args()

GENERIC_DOMAIN = BASE_GENERIC_DOMAIN + args.additional_generic_domain.split(args.list_separator)

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
    print(f'Completed reading {len(cbx_data)} employees.')
    print('Reading hiring client data file...')
    with open(hc_file, 'r', encoding=args.hc_encoding) as hc:
        for row in csv.reader(hc):
            hc_data.append(row)
    print(f'Completed reading {len(hc_data)} employees.')
    with open(output_file, 'w', newline='', encoding=args.output_encoding) as resultfile:
        writer = csv.writer(resultfile)
        headers = hc_data.pop(0)
        #todo check headers names
        headers.extend(['cbx_id', 'cbx_contractor', 'cbx_street', 'cbx_city', 'cbx_state', 'cbx_zip', 'cbx_country',
                        'expiration_date', 'registration_status', 'suspended', 'cbx_email',
                        'cbx_first_name', 'cbx_last_name', 'modules', 'cbx_account_type',
                        'cbx_subscription_fee', 'cbx_employee_price', 'parents', 'previous',
                        'hiring_clients', 'hiring_client_count', 'ratio_company', 'ratio_address', 'overall_score',
                        'contact_match', 'generic_domain', 'match_count', 'match_count_with_hc',
                        'multiple_company_match', 'analysis', 'index'])
        writer.writerow(headers)

        # match
        total = len(hc_data)
        index = 1
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
            if hc_row[HC_DO_NOT_MATCH].lower().strip() != 'true' or hc_row[HC_DO_NOT_MATCH].strip() != '1':
                if hc_force_cbx:
                    cbx_row = next(filter(lambda x: x[CBX_ID].strip() == hc_force_cbx, cbx_data), None)
                    hc_count = len(cbx_row[CBX_HIRING_CLIENTS].split(args.list_separator)) if cbx_row[CBX_HIRING_CLIENTS] else 0
                    matches.append({'cbx_id': cbx_row[CBX_ID],
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
                                'parents': cbx_row[CBX_PARENTS],
                                'previous': cbx_row[CBX_COMPANY_OLD],
                                'hiring_clients': cbx_row[CBX_HIRING_CLIENTS],
                                'hiring_client_count': hc_count,
                                'ratio_company': '',
                                'ratio_address': '',
                                'overall_score': '',
                                'contact_match': '',
                                })
                    continue
                for cbx_row in cbx_data:
                    cbx_email = cbx_row[CBX_EMAIL].lower()
                    cbx_domain = cbx_email[cbx_email.find('@') + 1:]
                    contact_match = False
                    if hc_domain in GENERIC_DOMAIN:
                        contact_match = True if cbx_email == hc_email else False
                    else:
                        contact_match = True if cbx_domain == hc_domain else False
                    cbx_zip = cbx_row[CBX_ZIP].replace(' ', '').upper()
                    cbx_company_en = cbx_row[CBX_COMPANY_EN]
                    cbx_company_en = re.sub(r"\([^()]*\)", "", cbx_company_en)
                    cbx_company_fr = cbx_row[CBX_COMPANY_FR]
                    cbx_company_fr = re.sub(r"\([^()]*\)", "", cbx_company_fr)
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

                    ratio_address = ratio_address + ratio_zip / 2
                    ratio_company = ratio_company_fr if ratio_company_fr > ratio_company_en else ratio_company_en
                    ratio_previous = 0
                    for item in cbx_previous.split(args.list_separator):
                        if item in (cbx_company_en, cbx_company_fr):
                            continue
                        ratio = fuzz.token_sort_ratio(item.lower().replace('.', '').replace(',', '').strip(),
                                                      clean_hc_company)
                        ratio_previous = ratio if ratio > ratio_previous else ratio_previous

                    ratio_previous *= 0.9
                    ratio_company = ratio_previous if ratio_previous > ratio_company else ratio_company
                    if ((ratio_company >= float(args.ratio_company) and ratio_address >= float(args.ratio_address)) or
                            contact_match):

                        overall_score = ratio_company
                        if cbx_row[CBX_REGISTRATION_STATUS] in ('Suspended' or 'Non Member') or not cbx_row[CBX_HIRING_CLIENTS]:
                            overall_score *= 0.8
                        cbx_company = cbx_company_fr if cbx_company_fr else cbx_company_en
                        print('   --> ', cbx_company, hc_email, cbx_row[CBX_ID], ratio_company, ratio_address, overall_score, contact_match)
                        parent_str = f'[parent: {cbx_parents}]' if cbx_parents else None
                        previous_str = f'[previous: {cbx_previous}]' if cbx_previous else None
                        display = [cbx_company]
                        cbx_company = ' '.join(display)
                        hc_count = len(cbx_row[CBX_HIRING_CLIENTS].split(args.list_separator)) if cbx_row[CBX_HIRING_CLIENTS] else 0
                        matches.append({'cbx_id': cbx_row[CBX_ID],
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
                                        'parents': cbx_row[CBX_PARENTS],
                                        'previous': cbx_row[CBX_COMPANY_OLD],
                                        'hiring_clients': cbx_row[CBX_HIRING_CLIENTS],
                                        'hiring_client_count': hc_count,
                                        'ratio_company': ratio_company,
                                        'ratio_address': ratio_address,
                                        'overall_score': overall_score,
                                        'contact_match': str(contact_match),
                                        })
            ids = []
            best_match = 0
            matches = sorted(matches, key=lambda x: (x['hiring_client_count'], x["overall_score"]), reverse=True)
            for item in matches[0:10]:
                ids.append(f'{item["cbx_id"]}, {item["company"]}, {item["address"]}, {item["zip"]}, {item["email"]}'
                           f' --> CR{item["ratio_company"]}, AR{item["ratio_address"]}, OS{item["overall_score"]},'
                           f' CM{item["contact_match"]}, HCC{item["hiring_client_count"]}')
            # append matching results to the hc_list
            uniques_cbx_id = set(item['cbx_id'] for item in matches)
            hc_row.append(matches[0]["cbx_id"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["company"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["address"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["city"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["state"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["zip"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["country"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["expiration_date"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["registration_status"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["suspended"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["email"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["first_name"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["last_name"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["modules"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["account_type"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["subscription_price"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["employee_price"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["parents"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["previous"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["hiring_clients"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["hiring_client_count"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["ratio_company"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["ratio_address"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["overall_score"] if len(uniques_cbx_id) else '')
            hc_row.append(matches[0]["contact_match"] if len(uniques_cbx_id) else '')
            hc_row.append(str(True if hc_domain in GENERIC_DOMAIN else False))
            hc_row.append(len(uniques_cbx_id) if len(uniques_cbx_id) else '')
            hc_row.append(str(len([i for i in matches if i['hiring_client_count'] > 0])))
            hc_row.append('True' if len(uniques_cbx_id) > 1 else 'False')
            hc_row.append('|'.join(ids))
            hc_row.append(str(index))
            writer.writerow(hc_row)
            print(f'{index} of {total} [{len(uniques_cbx_id)} found]')
            index += 1
