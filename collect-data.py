#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# articlestats.xls

# start sheet 'referencesamples'

# worksheet.cell(7, 4).value == 'article ID'
# articleid ids [8,4] (or reverse) for full text for sheet_names

import xlrd
import sys

class DupHolder:
    def __init__(self, placename):
        self.placename = placename
        self.dict_id_measure = {}
        self.has_dup = False
        self.num_dups = 0
        self.nums_ok = 0
        self.nums_most = 0
        self.nums_half = 0
        self.nums_less = 0
        self.nums_none = 0

class Gazetteer:
    def __init__(self, name_id,name,status,feat_type,nzgb_ref,land_district,crd_projection,crd_north,crd_east,crd_datum,crd_latitude,crd_longitude, accuracy, accuracy_rating):
        self.name_id = name_id
        self.name = name
        self.status = status
        self.feat_type = feat_type
        self.nzgb_ref = nzgb_ref
        self.land_district = land_district
        self.crd_projection = crd_projection
        self.crd_north = crd_north
        self.crd_east = crd_east
        self.crd_datum = crd_datum
        self.crd_latitude = crd_latitude
        self.crd_longitude = crd_longitude
        self.accuracy = accuracy
        self.accuracy_rating = accuracy_rating
        self.num_studies = 0

workbook = xlrd.open_workbook('articlestats.xls')
# workbook = xlrd.open_workbook('articlestats.xls', encoding='cp1252')
# workbook = xlrd.open_workbook('articlestats.xls', on_demand = True)

worksheet = workbook.sheet_by_name('referencesamples')
# worksheet = workbook.sheet_by_index(0)
# workbook.nsheets
# workbook.sheet_names()

list_OK = ('OK', 'Ok', 'ok')
list_MOST = ('MOST', 'Most', 'most')
list_HALF = ('HALF', 'Half', 'half')
list_LESS = ('LESS', 'Less', 'less')
list_NONE = ('NONE', 'None', 'none')
list_no_matches_at_all = ('no matches')

map_gazetteer = {}

gazbook = xlrd.open_workbook('gaz_names.xls')
gazsheet = gazbook.sheet_by_name('gaz_names')

for i in range(1, 50547):
    name_id = str(int(gazsheet.cell(i, 0).value))
    name = str(gazsheet.cell(i, 1).value)
    status = str(gazsheet.cell(i, 2).value)
    feat_type = str(gazsheet.cell(i, 3).value)
    nzgb_ref = str(gazsheet.cell(i, 4).value)
    land_district = str(gazsheet.cell(i, 5).value)
    crd_projection = str(gazsheet.cell(i, 6).value)
    crd_north = str(gazsheet.cell(i, 7).value)
    crd_east = str(gazsheet.cell(i, 8).value)
    crd_datum = str(gazsheet.cell(i, 9).value)
    crd_latitude = str(gazsheet.cell(i, 10).value)
    crd_longitude = str(gazsheet.cell(i, 11).value)
    accuracy = str(gazsheet.cell(i, 24).value)
    accuracy_rating = str(gazsheet.cell(i, 33).value)
    gaz = Gazetteer(name_id,name,status,feat_type,nzgb_ref,land_district,crd_projection,crd_north,crd_east,crd_datum,crd_latitude,crd_longitude,accuracy,accuracy_rating)
    map_gazetteer[name_id] = gaz
    # print("inserted map_gazetteer[{}] = {}".format(name_id, name))

print("articleid,all,OK,MOST,HALF,LESS,NONE")

for i in range(8, 297):
    if worksheet.cell(i, 4).value == xlrd.empty_cell.value:
        pass
    else:
        artcal = int(worksheet.cell(i, 4).value)
        # print(' row(+1) ' + str(i) + ' articleid ' + str(artcal))

        try:
            nshe = workbook.sheet_by_name(str(artcal))
            # print('sheet header ' + str(int(nshe.cell(0, 0).value)))
            artid = int(nshe.cell(0, 0).value)
            # check if cell 0 0 id is equal to artcal
            if artcal == artid:
                # print(str(artcal) + " " + str(artid))
                resultmap = {}
                rowcheck = 2
                okmatches = 0
                firstplace = str(nshe.cell(rowcheck, 0).value)
                firstmeasure = str(nshe.cell(rowcheck, 1).value)
                trynext = True
                sheetdict = {'OK': 0, 'MOST': 0, 'HALF': 0, 'LESS':0, 'NONE':0}

                while trynext:
                    try:
                        firstplace = str(nshe.cell(rowcheck, 0).value)
                        firstmeasure = str(nshe.cell(rowcheck, 1).value)

                        stripped1 = firstplace.strip(',')
                        stripped2 = stripped1.strip(')')
                        splitted1 = stripped2.split('(')
                        # print("split first {} second {}".format(splitted1[0].strip(), splitted1[1]))
                        placename = splitted1[0].strip()
                        placeid = splitted1[1]
                        gazElem = Gazetteer("name_id","name","status","feat_type","nzgb_ref","land_district","crd_projection","crd_north","crd_east","crd_datum","crd_latitude","crd_longitude", "accuracy", "accuracy_rating")
                        try:
                            gazElem = map_gazetteer[str(placeid)]
                        except KeyError:
                            print("not found placeid: {} name: {}".format(placeid, placename))

                        nameobj = DupHolder(placename)

                        try:
                            workobj = resultmap[placename]
                            nameobj = workobj
                            nameobj.has_dup = True
                            nameobj.num_dups = nameobj.num_dups + 1

                        except KeyError:
                            # print("not yet objectified: {}".format(placename))
                            resultmap[placename] = nameobj

                        if firstmeasure in list_OK:
                            sheetdict['OK'] = sheetdict['OK'] + 1
                            nameobj.nums_ok = nameobj.nums_ok + 1
                            gazElem.num_studies = gazElem.num_studies + 1

                        if firstmeasure in list_MOST:
                            sheetdict['MOST'] = sheetdict['MOST'] + 1
                            nameobj.nums_most = nameobj.nums_most + 1
                            gazElem.num_studies = gazElem.num_studies + 1

                        if firstmeasure in list_HALF:
                            sheetdict['HALF'] = sheetdict['HALF'] + 1
                            nameobj.nums_half = nameobj.nums_half + 1
                            # gazElem.num_studies = gazElem.num_studies + 1

                        if firstmeasure in list_LESS:
                            sheetdict['LESS'] = sheetdict['LESS'] + 1
                            nameobj.nums_less = nameobj.nums_less + 1

                        if firstmeasure in list_NONE:
                            sheetdict['NONE'] = sheetdict['NONE'] + 1
                            nameobj.nums_none = nameobj.nums_none + 1

                        rowcheck = rowcheck+1

                    except:
                        print("hit empty on: " + str(rowcheck))
                        trynext = False
                        pass

                actual_rows = rowcheck-2
                print("found {} rows in sheet {}".format(actual_rows, artid))
                print("{},{},{},{},{},{},{}".format(artid,actual_rows,sheetdict['OK'],sheetdict['MOST'],sheetdict['HALF'],sheetdict['LESS'],sheetdict['NONE']))
                for elemkey in resultmap:
                    placeobj = resultmap[elemkey]
                    print("placeobj, {}, nums_ok, {}, nums_most, {}, nums_half, {}, nums_less, {}, nums_none, {}, has_dupl, {}, dup_count, {}".format(placeobj.placename, placeobj.nums_ok, placeobj.nums_most, placeobj.nums_half, placeobj.nums_less, placeobj.nums_none, placeobj.has_dup, str(placeobj.num_dups)))

                # print("{} NONEs that have #{} same / duplicate name with OK")

            else:
                print("not equal: " + str(artcal) + " " + str(artid))
        except:
            print("Unexpected error:", sys.exc_info()[0])
            print('error couldnt retrieve sheet no ' + str(artcal))
            pass

print("len gaz = {} ".format(len(map_gazetteer)))

colHeader = """{
    "type": "FeatureCollection",
    "features": ["""

colFooter = """    ]
}"""

with open('geoJson.json', 'a') as out:
    out.write(colHeader)
    for each_gaz_key in map_gazetteer:
        each_gaz = map_gazetteer[each_gaz_key]
        if each_gaz.num_studies > 0:
            props_arr = """ "name_id": "{}",
                    "name": "{}",
                    "status": "{}",
                    "feat_type": "{}",
                    "land_district": "{}",
                    "num_studies": "{}"
    """.format(each_gaz.name_id, each_gaz.name, each_gaz.status, each_gaz.feat_type, each_gaz.land_district, each_gaz.num_studies)

            json = """{
                "type": "Feature",
                "geometry": {
        "type": "Point",
        "coordinates": [
            """+each_gaz.crd_longitude+""",
            """+each_gaz.crd_latitude+"""
        ]
    },
                "properties": {
                    """ + props_arr + """
                }
            }"""
            out.write(json + ',\n')
        else:
            pass

    out.write(colFooter + '\n')
    out.flush()
    out.close()

# name_id,name,status,feat_type,nzgb_ref,land_district
# eventually print/write csv or so with just articleid and flags for title, abs, full
# and then csv again for each journal
# stat anal anova whatever
