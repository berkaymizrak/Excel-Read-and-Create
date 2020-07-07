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
Connect.check_run(code, program, 30, sound_error=True)  # <-- Remove this line in your app or you can create yours.


excel_read = input('\nWhat is your excel name which will be read? (For default, leave empty - Import.xlsx): ')
if not excel_read:
    excel_read = 'Import.xlsx'

excel_create = input('\nWhat is your excel name which will be created? (For default, leave empty - Export.xlsx): ')
if not excel_create:
    excel_create = 'Export.xlsx'

print()

# //////// EXCEL READ \\\\\\\\
my_excel_data = File.excel_read_to_dict(excel_read)  # <-- HERE
# it returns a dictionary from 3 rows excel file as:
# my_excel_data = {
#     1: ['1st Column Value', '2nd Column Value', '3rd Column Value', '4th Column Value', '5th Column Value', ],
#     2: ['1st Column Value', '2nd Column Value', '3rd Column Value', '4th Column Value', '5th Column Value', ],
#     3: ['1st Column Value', '2nd Column Value', '3rd Column Value', '4th Column Value', '5th Column Value', ],
# }

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
# Headers and Sizes are optional.
# File.excel_create(excel_create, my_excel_data)  # <-- HERE
File.excel_create(excel_create, my_excel_data, headers=headers, sizes=sizes, page_name='Sales Orders New')  # <-- HERE


message = 'Program successfully worked.'
Progress.exit_app(message=message, exit_all=False)

Progress.count_down(60, message='Shot down...')
