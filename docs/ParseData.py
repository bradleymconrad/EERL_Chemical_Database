# -*- coding: utf-8 -*-
from openpyxl import load_workbook

# Get workbook and worksheet
wb = load_workbook(filename='..\\Chemical_info.xlsx', read_only=True)
ws = wb['Chemical_Tracking']

GroupNames = {'Hydrocarbon', 'Acid', 'Alcohol', 'Other'}

for Group in GroupNames:

    # Initiate data
    names = []
    locs1 = []
    locs2 = []
    vess1 = []
    vess2 = []
    cont1 = []
    cont2 = []
    cont3 = []
    notes = []

    # Pull data
    for row in ws.rows:
        # Skip header data
        if row[0].value in ('Chemical Name', "Tim Horton's Dark Roast"):
            continue

        # Exit upon empty row
        if row[0].value is None:
            break

        if row[1].value != Group:
            continue

        names.append(row[0].value)
        locs1.append(row[2].value)
        locs2.append(row[3].value)
        vess1.append(row[4].value)
        cont1.append(row[5].value)
        cont2.append(row[6].value)
        cont3.append(row[7].value)
        notes.append(row[8].value)

    # Create group .rst document
    workfile = f'support\\{Group}s.rst'
    with open(workfile, 'w') as f:
        # Print header
        f.write('#' * (len(Group) + 1))
        f.write(f'\n{Group}s\n')
        f.write('#' * (len(Group) + 1))
        f.write('\n\n')

        # Print toc
        f.write('.. contents::\n  :depth: 1\n  :local:\n')

        # Print chemical data!
        for index, _ in enumerate(names):
            f.write('\n')

            if index > 0:
                if names[index] != names[index - 1]:
                    f.write('*' * len(names[index]))
                    f.write(f'\n{names[index]}\n')
                    f.write('*' * len(names[index]))
            else:
                f.write('*' * len(names[index]))
                f.write(f'\n{names[index]}\n')
                f.write('*' * len(names[index]))

            f.write('\n\n.. list-table::\n  :widths: 25 75\n  :align: center\n\n')

            f.write(f'  * - Location:\n    - | {locs1[index]}\n      | {locs2[index]}\n')

            f.write(f'  * - Vessel:\n    - | {vess1[index]}\n')

            f.write(f'  * - | Emergency\n      | Contact:\n    - | {cont1[index]}\n      | {cont2[index]}\n      | {cont3[index]}\n')

            if notes[index] is not None:
                f.write(f'  * - Notes:\n    - {notes[index]}\n')

            f.write(f'  * - MSDS:\n    - :download:`pdf <../../SDS_repository/{names[index]}.pdf>`\n')
