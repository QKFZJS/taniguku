# coding: utf-8
from excel import ExcelBook, ExcelSheet
from schema import Schema
import copy

COLOR_LIGHT = 'DDDDDD'
COLOR_BLACK = '000000'
COLOR_WEEK_BLACK = '222222'

COLOR_BG_DARK = '808080'

schema = Schema()
schema.load("./ansible-playbook")

excel = ExcelBook("./result.xlsx")

def addBaseBorder(target, rng, endpoint, additionalBorder):
    target.addBorder(list(rng), [endpoint], 'thin', 
                     COLOR_WEEK_BLACK, ['top'])
    target.addBorder(['B','I','J',rng[0:1]], [endpoint], 
                    'thin', COLOR_WEEK_BLACK, ['left'])
    target.addBorder(additionalBorder, [endpoint], 
                    'thin', COLOR_WEEK_BLACK, ['left'])
    target.addBorder(['J'], [endpoint], 'thin', 
                    COLOR_WEEK_BLACK, ['right'])
    target.addFont(list("BCDEFGHIJ"), [endpoint], name='Arial')
 

# Value tree to excel.
#
def dictValueSheetnize(target, src, rng, endpoint, additionalBorder=[]):
    for k in src.keys():
        if isinstance(src[k], (list, dict)):
            target.addData(rng[0:1], endpoint, k)
            addBaseBorder(target, rng, endpoint, additionalBorder)
            endpoint += 1
            ab = copy.deepcopy(additionalBorder) 
            ab.append(rng[0:1])

            if isinstance(src[k], list):
                endpoint = listValueSheetnize(target, src[k], rng[1:], 
                                              endpoint, ab)
            else:
                endpoint = dictValueSheetnize(target, src[k], rng[1:], 
                                              endpoint, ab)
 
            target.addBorder(list(rng), [endpoint], 'thin', 
                             COLOR_WEEK_BLACK, ['buttom'])
        elif isinstance(src[k], (int, float, str, bool)):
            target.addData(rng[0:1], endpoint, k)
            if isinstance(src[k], bool):
                target.addData('I', endpoint, "yes" if src[k] else "no")
            else:
                target.addData('I', endpoint, src[k])

            addBaseBorder(target, rng, endpoint, additionalBorder)
            target.addBorder(list(rng), [endpoint], 'thin', 
                             COLOR_WEEK_BLACK, ['buttom'])
            endpoint += 1
    return endpoint

def listValueSheetnize(target, src, rng, endpoint, additionalBorder=[]):
    i = 0
    for item in src:
        if isinstance(item, (list, dict)):
            target.addData(rng[0:1], endpoint, '[ '+str(i)+' ]')
            addBaseBorder(target, rng, endpoint, additionalBorder)
            endpoint += 1
            ab = copy.deepcopy(additionalBorder) 
            ab.append(rng[0:1])

            if isinstance(item, list):
                endpoint = listValueSheetnize(target, item, rng[1:], 
                                              endpoint, ab)
            else:
                endpoint = dictValueSheetnize(target, item, rng[1:], 
                                              endpoint, ab)

            target.addBorder(list(rng), [endpoint], 'thin', 
                             COLOR_WEEK_BLACK, ['buttom'])

        elif isinstance(item, (int, float, str, bool)):
            target.addData(rng[0:1], endpoint, '[ '+str(i)+' ]')
            if isinstance(item, bool):
                target.addData('I', endpoint, "yes" if item else "no")
            else:
                target.addData('I', endpoint, item)

            addBaseBorder(target, rng, endpoint, additionalBorder)
            target.addBorder(list(rng), [endpoint], 'thin', 
                             COLOR_WEEK_BLACK, ['buttom'])
            endpoint += 1

        i += 1
    return endpoint

def valueSheetnize(target, src, rng, endpoint, additionalBorder=[]):
    if isinstance(src, dict):
        endpoint = dictValueSheetnize(target, src, rng, endpoint, additionalBorder)
    elif isinstance(src, list):
        endpoint = listValueSheetnize(target, src, rng, endpoint, additionalBorder)
    return endpoint

# Create Site Sheet.
#
sheet = excel.createExcelSheet(0, "site")
sheet.addRowDimension('B', 30)
sheet.addRowDimension('C', 30)
sheet.addRowDimension('D', 30)

sheet.addData('B', 1, "site.yml")
sheet.addData('B', 2, "設定")
sheet.addData('C', 2, "値")
sheet.addData('B', 3, "import_playbook")

sheet.addFont(['B','C'], [1], name='Arial', bold=True)
sheet.addFont(['B','C'], [2], name='Meiryo UI', color=COLOR_LIGHT)

sheet.addBackgroundColor(['B','C'], [2], 'solid', COLOR_BG_DARK)

sheet.addBorder(['B','C'], [2], 'thin', COLOR_WEEK_BLACK, 
                ['top','bottom','left','right'])
endpoint = 3
for playbook in schema.getSite():
    sheet.addData('C', endpoint, playbook)
    sheet.addBorder(['B','C'], [endpoint], 'thin', COLOR_WEEK_BLACK,
                    ['left','right'])
    sheet.addBorder(['C'], [endpoint], 'dotted', COLOR_BLACK,
                    ['bottom'])
    sheet.addFont(['B','C'], [endpoint], name='Arial')
    endpoint += 1
sheet.addBorder(['B','C'], [endpoint-1], 'thin', COLOR_BLACK,
                ['bottom'])

sheet.buildSheet()

# Create Playbook Sheet
#
sheet = excel.createExcelSheet(1, "playbook")
sheet.addRowDimension('B', 4)
sheet.addRowDimension('C', 4)
sheet.addRowDimension('D', 4)
sheet.addRowDimension('E', 4)
sheet.addRowDimension('F', 4)
sheet.addRowDimension('G', 4)
sheet.addRowDimension('H', 20)
sheet.addRowDimension('I', 70)
sheet.addRowDimension('J', 50)

sheet.addData('B', 1, "playbook.yml")
sheet.addData('B', 2, "設定")
sheet.addData('I', 2, "値")
sheet.addData('J', 2, "備考")

sheet.addFont(['B'], [1], name='Arial', bold=True)
sheet.addFont(['B','I','J'], [2], name='Meiryo UI', color=COLOR_LIGHT)

sheet.addBackgroundColor(list("BCDEFGHIJ"), [2], 'solid', COLOR_BG_DARK)

sheet.addBorder(list("BCDEFGHIJ"), [2], 'thin', COLOR_WEEK_BLACK, 
                ['top','bottom'])
sheet.addBorder(['B','I','J'], [2], 'thin', COLOR_WEEK_BLACK, 
                ['left'])
sheet.addBorder(['J'], [2], 'thin', COLOR_WEEK_BLACK, 
                ['right'])

endpoint = 3
for playbook in schema.getPlaybooks():
    sheet.addData('B', endpoint, playbook['path'])
    sheet.addFont(list("BCDEFGHIJ"), [endpoint], name='Arial')
    sheet.addBorder(list("BCDEFGHIJ"), [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['top'])
    sheet.addBorder(['B','I','J'], [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['left'])
    sheet.addBorder(['J'], [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['right'])
    endpoint += 1

    sheet.addData('C', endpoint, 'hosts')
    sheet.addData('I', endpoint, playbook['hosts'])
    sheet.addData('C', endpoint+1, 'become')
    sheet.addData('I', endpoint+1, 
                  "yes" if playbook['become'] else "no")
    sheet.addBorder(list("CDEFGHIJ"), [endpoint, endpoint+1], 'thin', 
                    COLOR_WEEK_BLACK, ['top','bottom'])
    sheet.addBorder(list("BCIJ"), [endpoint, endpoint+1], 'thin', 
                    COLOR_WEEK_BLACK, ['left'])
    sheet.addBorder(['J'], [endpoint, endpoint+1], 'thin', 
                    COLOR_WEEK_BLACK, ['right'])
    sheet.addFont(list("BCDEFGHIJ"), [endpoint, endpoint+1], name='Arial')
    endpoint += 2

    sheet.addData('C', endpoint, 'roles')
    for role in playbook['roles']:
        sheet.addData('I', endpoint, role)
        sheet.addBorder(list("IJ"), [endpoint], 'dotted', 
                        COLOR_WEEK_BLACK, ['bottom'])
        sheet.addBorder(list("BCIJ"), [endpoint], 'thin', 
                        COLOR_WEEK_BLACK, ['left'])
        sheet.addBorder(['J'], [endpoint], 'thin', 
                        COLOR_WEEK_BLACK, ['right'])
        sheet.addFont(list("BCDEFGHIJ"), [endpoint], name='Arial')
        endpoint += 1
    sheet.addBorder(list("BCDEFGHIJ"), [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['top'])

sheet.buildSheet()



# Create Global vars Sheet
#
sheet = excel.createExcelSheet(2, "global vars")
sheet.addRowDimension('B', 4)
sheet.addRowDimension('C', 4)
sheet.addRowDimension('D', 4)
sheet.addRowDimension('E', 4)
sheet.addRowDimension('F', 4)
sheet.addRowDimension('G', 4)
sheet.addRowDimension('H', 20)
sheet.addRowDimension('I', 70)
sheet.addRowDimension('J', 50)

sheet.addData('B', 1, "group_vars/all.yml")
sheet.addData('B', 2, "設定")
sheet.addData('I', 2, "値")
sheet.addData('J', 2, "備考")

sheet.addFont(['B'], [1], name='Arial', bold=True)
sheet.addFont(['B','I','J'], [2], name='Meiryo UI', color=COLOR_LIGHT)

sheet.addBackgroundColor(list("BCDEFGHIJ"), [2], 'solid', COLOR_BG_DARK)

sheet.addBorder(list("BCDEFGHIJ"), [2], 'thin', COLOR_WEEK_BLACK, 
                ['top','bottom'])
sheet.addBorder(['B','I','J'], [2], 'thin', COLOR_WEEK_BLACK, 
                ['left'])
sheet.addBorder(['J'], [2], 'thin', COLOR_WEEK_BLACK, 
                ['right'])

endpoint = valueSheetnize(sheet, schema.getGlobalVars(), "BCDEFGHIJ", 3)            

sheet.addBorder(list("BCDEFGHIJ"), [endpoint], 'thin', COLOR_WEEK_BLACK, 
                ['top'])

sheet.buildSheet()

# group sheet
#
sheet = excel.createExcelSheet(3, "group")
sheet.addRowDimension('B', 4)
sheet.addRowDimension('C', 4)
sheet.addRowDimension('D', 4)
sheet.addRowDimension('E', 4)
sheet.addRowDimension('F', 4)
sheet.addRowDimension('G', 4)
sheet.addRowDimension('H', 20)
sheet.addRowDimension('I', 70)
sheet.addRowDimension('J', 50)

sheet.addData('B', 1, "group")
sheet.addData('B', 2, "設定")
sheet.addData('I', 2, "値")
sheet.addData('J', 2, "備考")

sheet.addFont(['B'], [1], name='Arial', bold=True)
sheet.addFont(['B','I','J'], [2], name='Meiryo UI', color=COLOR_LIGHT)

sheet.addBackgroundColor(list("BCDEFGHIJ"), [2], 'solid', COLOR_BG_DARK)

sheet.addBorder(list("BCDEFGHIJ"), [2], 'thin', COLOR_WEEK_BLACK, 
                ['top','bottom'])
sheet.addBorder(['B','I','J'], [2], 'thin', COLOR_WEEK_BLACK, 
                ['left'])
sheet.addBorder(['J'], [2], 'thin', COLOR_WEEK_BLACK, 
                ['right'])

endpoint = 3
for group in schema.getGroups():
    sheet.addData('B', endpoint, group['name'])
    sheet.addFont(list("BCDEFGHIJ"), [endpoint], name='Arial')
    sheet.addBorder(list("BCDEFGHIJ"), [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['top'])
    sheet.addBorder(['B','I','J'], [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['left'])
    sheet.addBorder(['J'], [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['right'])
    endpoint += 1
 
    sheet.addData('C', endpoint, 'hosts')
    sheet.addBorder(list("CDEFGHIJ"), [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['top'])
    for host in group['hosts']:
        sheet.addData('I', endpoint, host)
        sheet.addBorder(list("IJ"), [endpoint], 'dotted', 
                        COLOR_WEEK_BLACK, ['bottom'])
        sheet.addBorder(list("BCIJ"), [endpoint], 'thin', 
                        COLOR_WEEK_BLACK, ['left'])
        sheet.addBorder(['J'], [endpoint], 'thin', 
                        COLOR_WEEK_BLACK, ['right'])
        sheet.addFont(list("BCDEFGHIJ"), [endpoint], name='Arial')
        endpoint += 1
    sheet.addData('C', endpoint, 'vars')
    sheet.addFont(list("BCDEFGHIJ"), [endpoint], name='Arial')
    sheet.addBorder(list("CDEFGHIJ"), [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['top'])
    sheet.addBorder(['B','C','I','J'], [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['left'])
    sheet.addBorder(['J'], [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['right'])
 
    endpoint += 1

    endpoint = valueSheetnize(sheet, group['vars'], "DEFGHIJ", endpoint, ['C'])            
     
sheet.addBorder(list("BCDEFGHIJ"), [endpoint], 'thin', COLOR_WEEK_BLACK, 
                ['top'])

sheet.buildSheet()

# host sheet
#
sheet = excel.createExcelSheet(4, "host")
sheet.addRowDimension('B', 4)
sheet.addRowDimension('C', 4)
sheet.addRowDimension('D', 4)
sheet.addRowDimension('E', 4)
sheet.addRowDimension('F', 4)
sheet.addRowDimension('G', 4)
sheet.addRowDimension('H', 20)
sheet.addRowDimension('I', 70)
sheet.addRowDimension('J', 50)

sheet.addData('B', 1, "group")
sheet.addData('B', 2, "設定")
sheet.addData('I', 2, "値")
sheet.addData('J', 2, "備考")

sheet.addFont(['B'], [1], name='Arial', bold=True)
sheet.addFont(['B','I','J'], [2], name='Meiryo UI', color=COLOR_LIGHT)

sheet.addBackgroundColor(list("BCDEFGHIJ"), [2], 'solid', COLOR_BG_DARK)

sheet.addBorder(list("BCDEFGHIJ"), [2], 'thin', COLOR_WEEK_BLACK, 
                ['top','bottom'])
sheet.addBorder(['B','I','J'], [2], 'thin', COLOR_WEEK_BLACK, 
                ['left'])
sheet.addBorder(['J'], [2], 'thin', COLOR_WEEK_BLACK, 
                ['right'])

endpoint = 3
for host in schema.getHosts():
    sheet.addData('B', endpoint, host['name'])
    sheet.addFont(list("BCDEFGHIJ"), [endpoint], name='Arial')
    sheet.addBorder(list("BCDEFGHIJ"), [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['top'])
    sheet.addBorder(['B','I','J'], [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['left'])
    sheet.addBorder(['J'], [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['right'])
    endpoint += 1
 
    sheet.addData('C', endpoint, 'vars')
    sheet.addFont(list("BCDEFGHIJ"), [endpoint], name='Arial')
    sheet.addBorder(list("CDEFGHIJ"), [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['top'])
    sheet.addBorder(['B','C','I','J'], [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['left'])
    sheet.addBorder(['J'], [endpoint], 'thin', COLOR_WEEK_BLACK, 
                    ['right'])
 
    endpoint += 1

    endpoint = valueSheetnize(sheet, host['vars'], "DEFGHIJ", endpoint, ['C'])            
     
sheet.addBorder(list("BCDEFGHIJ"), [endpoint], 'thin', COLOR_WEEK_BLACK, 
                ['top'])

sheet.buildSheet()


excel.save()