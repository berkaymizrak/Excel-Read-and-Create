# -*- coding: utf-8 -*-
# Excel Read and Create
# Berkay MIZRAK
# www.BerkayMizrak.com
# www.DaktiNetwork.com

version = '1.2'
program = "Excel Read and Create v%s" % version
code = 'excel_read_create'

print('\n\t%s' % (program))
print('\n\t\twww.BerkayMizrak.com')
print('\n\t\t\twww.DaktiNetwork.com')

try:
    from Functions import Connect
    from Functions import File
    from Functions import Progress
except Exception as e:
    print()
    print(e)
    while True:
        input('\n! ! ERROR --> A module is not installed...')


# Check if program has permission to run from developer by API
# Connect.check_run(code, program, 30, sound_error=True)  # <-- Remove this line in your app or you can create yours.

excel_read = input('\nWhat is your excel name which will be read? (For default, leave empty - Import.xlsx): ')
if not excel_read:
    excel_read = 'Import.xlsx'

excel_create_file = input('\nWhat is your excel name which will be created? (For default, leave empty - Export.xlsx): ')
if not excel_create_file:
    excel_create_file = 'Export.xlsx'

print()

# //////// EXCEL READ \\\\\\\\
my_excel_data, excel_headers = File.excel_read_to_dict(excel_read)  # <-- HERE
# it returns a dictionary from 3 rows excel file as:
"""
my_excel_data = {
    1: ['1st Column Value', '2nd Column Value', '3rd Column Value', '4th Column Value', '5th Column Value', ],
    2: ['1st Column Value', '2nd Column Value', '3rd Column Value', '4th Column Value', '5th Column Value', ],
    3: ['1st Column Value', '2nd Column Value', '3rd Column Value', '4th Column Value', '5th Column Value', ],
}
"""
"""
excel_headers = {
    1: ['ID', 'OrderDate', 'Region', 'Rep', 'Item', 'Units', 'Unit Cost', 'Total'],
    2: ['Header 1', 'Header 2', 'Header 3', 'Header 4', ]  # <-- If there is second or more header rows!
}
"""

print()

# //////// EXCEL CREATE \\\\\\\\
# You can define headers, sizes and page_name. For instance:
headers = [
    'ID from Import',
    'OrderDate',
    'Region',
    'Rep',
    'Item',
    'Units',
    'Unit Cost',
    'Total',
]
sizes = [
    13.5,
    10,
    21,
    21,
    17,
    13,
    13,
    13,
]
# Sizes list defines each column's width. You can leave empty to use default width.

AdditionalAttributes_1 = {
    'My Header 1': 'Additional value 1',
    'My Header 2': 'Additional value 2',
    'My Header 3': 'Additional value 3',
}
AdditionalAttributes_2 = {
    'My Header 3': 'Additional value 3',
    'My Header 1': 'Additional value 1',
    'My Header 2': 'Additional value 2',
}
my_excel_data[1].append(AdditionalAttributes_1)
my_excel_data[2].append(AdditionalAttributes_2)
# Additional attributes are putting values to the cells which is under the same HEADERS.
# In this way you can push mixed dictionary but they will be under same headers based on DICTIONARY KEY.

# Additional Attributes, Headers and Sizes are optional.

# Option 1 - Push excel data without defining sizes, headers etc.
File.excel_create(excel_create_file, my_excel_data)  # <-- HERE - Option 1
# Option 2 - Push excel data with defining sizes, headers etc. + Use headers from the list defined above.
File.excel_create(excel_create_file, my_excel_data, headers=headers, sizes=sizes, page_name='Sales Orders New')  # <-- HERE - Option 2
# Option 3 - Use headers from reading excel.
File.excel_create(excel_create_file, my_excel_data, headers=excel_headers[1], sizes=sizes, page_name='Sales Orders New')  # <-- HERE - Option 3


# //////// Done. Wait and shot down. \\\\\\\\
message = 'Program successfully worked.'
Progress.exit_app(message=message, exit_all=False)

Progress.count_down(60, message='Shot down...')
exit()

