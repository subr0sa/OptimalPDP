from urllib.parse import quote
from urllib import request
import json
import itertools
import xlrd
from openpyxl import Workbook
import random



amap_web_key = 'daee36485ea08ec6e9871a4d4bcf5a46'
driving_url = 'https://restapi.amap.com/v3/direction/driving'
walking_url = 'https://restapi.amap.com/v3/direction/walking'


# ---------------------------------------------------  Function Part ---------------------------------------------------
def read_excel_to_dict(xlsxfile):
    data = xlrd.open_workbook(xlsxfile)
    sheet = data.sheet_by_index(0)
    all_dic = {}
    for i in range(1, sheet.nrows):
        key = str(sheet.row_values(i)[1])
        all = [sheet.row_values(i)[2],sheet.row_values(i)[3],sheet.row_values(i)[4],sheet.row_values(i)[5]]
        all_dic[key] = all
    return all_dic


def random_dr_altdic(all_dic):
    alt_key = random.sample(all_dic.keys(),6)
    alt_dic = dict([(key, all_dic[key]) for key in alt_key])
    dr_key = (random.sample(alt_key,1))[0]
    dr = alt_dic[dr_key][3]
    rd = [dr,alt_dic]
    return rd


def routingorder(dr,al):
    fastest = []
    ls = itertools.permutations(al, 6)
    AB = pufirst(ls, al[0], al[1])
    CD = pufirst(AB, al[2], al[3])
    EF = pufirst(CD, al[4], al[5])
    t = 43200 #12 hrs
    for i in EF:
        d_data = driving_routes(dr,i[0], i[1], i[2], i[3], i[4], i[5])
        ori_time = d_data['route']['paths'][0]['duration']
        time = int(ori_time)
        if time < t:
            t = time
            fastest = [dr,i[0],i[1],i[2],i[3],i[4],i[5]]
    return fastest


def driving_info(data):
    steps = data['route']['paths'][0]['steps']
    n = 0
    c = 0
    s = 0
    for step in steps:
        s = s+1
        if step["action"] == "左转" or step["action"] == "右转" or step["action"] == "左转调头":
            n = n + 1
        if step["tmcs"][0]["status"] == "畅通":
            c = c + 1
    driving_info = {"distance":data['route']['paths'][0]['distance'], "time":data['route']['paths'][0]['duration'],\
       "trafficlight":data['route']['paths'][0]['traffic_lights'], "fee":data['route']['taxi_cost'],"turns":str(n),"sections":str(s),"clear":str(c)}
    return driving_info


def driving_routes(ori_lonlat, wp1_lonlat, wp2_lonlat, wp3_lonlat, wp4_lonlat, wp5_lonlat, des_lonlat):
    ori_des = "&origin=" + str(ori_lonlat) + "&destination=" + str(des_lonlat)
    waypoints = "&waypoints=" + str(wp1_lonlat) + ";" + str(wp2_lonlat) + ";" + str(wp3_lonlat) + ";" + str(wp4_lonlat) + ";" \
                + str(wp5_lonlat) + ";"
    strategy = "&strategy=10" # strategy see https://lbs.amap.com/api/webservice/guide/api/direction#driving
    dreq_url = driving_url + "?key=" + amap_web_key + ori_des + waypoints + strategy + "&extensions=all&output=json"
    driving_data = ""
    with request.urlopen(dreq_url) as f:
        driving_data = f.read()
        driving_data = json.loads(driving_data)
    return driving_data


def driving_routes_each(ori_lonlat,des_lonlat):
    ori_des = "&origin=" + str(ori_lonlat) + "&destination=" + str(des_lonlat)
    strategy = "&strategy=10" # strategy see https://lbs.amap.com/api/webservice/guide/api/direction#driving
    dreq_url = driving_url + "?key=" + amap_web_key + ori_des + strategy + "&extensions=all&output=json"
    driving_data = ""
    with request.urlopen(dreq_url) as f:
        driving_data = f.read()
        driving_data = json.loads(driving_data)
    return driving_data


def walking_info(ori_lonlat, alt_lonlat):
    ori = "&origin=" + str(ori_lonlat)
    des = "&destination=" + str(alt_lonlat)
    wreq_url = walking_url + "?key=" + amap_web_key + ori + des
    walking_data = ""
    walking_info = {}
    with request.urlopen(wreq_url) as f:
        walking_data = f.read()
        walking_data = json.loads(walking_data)
        steps = walking_data["route"]["paths"][0]["steps"]
        cw = 0
        disa = 0
        for step in steps:
            if step["walk_type"] == 1:
                cw = cw +1
            if step["walk_type"] == 4 or step["walk_type"] == 8 or step["walk_type"] == 9 or step["walk_type"] == 20 or step["walk_type"] == 21:
                disa = disa + 1
        walking_info = {"distance":walking_data["route"]["paths"][0]["distance"],"time":walking_data["route"]["paths"][0]["duration"],\
                        "crosswalk":str(cw), "disabled":str(disa)}
    return walking_info


def pufirst(ls, a, b):
    ol = []
    for l in ls:
        t = list(l)
        if t[0] != b:
            if not(t[1] == b and t[2] == a):
                if not(t[1] == b and t[3] == a) and not(t[2] == b and t[3] == a):
                    if not(t[1] == b and t[4] == a) and not(t[2] == b and t[4] == a) and not(t[3] == b and t[4] == a):
                        if not (t[1] == b and t[5] == a) and not (t[2] == b and t[5] == a) and not (
                                t[3] == b and t[5] == a) and not (t[4] == b and t[5] == a):
                            ol.append(t)

    return ol

#oporder = oporder[1-6],
def walkingtime_list(oporder):
    wt_list = []
    for i in [0,1,2]:
        wt = walking_info(oporder,alt_dic[oporder][i])['time']
        wt_list.append(wt)
    return wt_list

#or = Automobile Coordinates, oporder = optimal order
def optimal_route(ori,oporder):
    alt_num = [0,1,2]
    time_min = 43200
    each_time = 0
    ijklmn = []
    route_info = {}
    coor = []
    coor_info = []
    route_data = []
    way1_wt = walkingtime_list(oporder[1])
    way2_wt = walkingtime_list(oporder[2])
    way3_wt = walkingtime_list(oporder[3])
    way4_wt = walkingtime_list(oporder[4])
    way5_wt = walkingtime_list(oporder[5])
    des_wt = walkingtime_list(oporder[6])
    for i in alt_num:
        way1 = alt_dic[oporder[1]][i]
        for j in alt_num:
            way2 = alt_dic[oporder[2]][j]
            for k in alt_num:
                way3 = alt_dic[oporder[3]][k]
                for l in alt_num:
                    way4 = alt_dic[oporder[4]][l]
                    for m in alt_num:
                        way5 = alt_dic[oporder[5]][m]
                        for n in alt_num:
                            des = alt_dic[oporder[6]][n]
                            route_data = driving_routes(ori,way1,way2,way3, way4, way5, des)
                            dri_time = int(driving_info(route_data)["time"])
                            walk_time = int(way1_wt[i]) + int(way2_wt[j]) + int(way3_wt[k]) + int(way4_wt[l]) + int(way5_wt[m]) + int(des_wt[n])
                            each_time = int(dri_time) + walk_time
                            time_min = int(time_min)
                            if each_time < time_min:
                                time_min = each_time
                                ijklmn = [i,j,k,l,m,n]
                                coor = [ori, way1, way2, way3, way4, way5, des]
    route_info = driving_info(route_data)
    coor_info = [route_info, coor, ijklmn]
    return coor_info


def all_info(oporder):
    ori_route1_data = driving_routes(oporder[0],oporder[1],oporder[2],oporder[3], oporder[4],oporder[5], oporder[6])
    ori_route1 = driving_info(ori_route1_data)
    op_route1 = optimal_route(dr,oporder)
    walk_all = []
    or_drive_all = []
    op_drive_all = []
    for i in [1,2,3,4,5,6]:
        j = i - 1
        walk_each = walking_info(oporder[i],op_route1[1][i])
        walk_all.append(walk_each)
        or_drive_data = driving_routes_each(oporder[j],oporder[i])
        or_drive_each = driving_info(or_drive_data)
        or_drive_all.append(or_drive_each)
        op_drive_data = driving_routes_each(op_route1[1][j],op_route1[1][i])
        op_drive_each = driving_info(op_drive_data)
        op_drive_all.append(op_drive_each)
    info = {"OriginalRoute Information":ori_route1,"OptimalRoute Information":op_route1,"Walking Information":walk_all,"OriDrivingEachSection":or_drive_all,"OptDrivingEachSection":op_drive_all}
    return info

def optimal_order_info(dr,alt_dic):
    info = {}
    oriCoor = list(alt_dic.keys())
    al = [oriCoor[0],oriCoor[1],oriCoor[2],oriCoor[3],oriCoor[4],oriCoor[5]]
    name_dict = {oriCoor[0]: "P1_P", oriCoor[1]: "P1_D", oriCoor[2]: "P2_P", oriCoor[3]: "P2_D", oriCoor[4]: "P3_P",
                 oriCoor[5]: "P3_D"}
    oporder = routingorder(dr,al)
    reordername_dict = [name_dict[oporder[1]], name_dict[oporder[2]], name_dict[oporder[3]], name_dict[oporder[4]],
                        name_dict[oporder[5]], name_dict[oporder[6]]]
    oporder_rename = [oporder, reordername_dict]
    return oporder_rename


def each_pas_info(pasOrder,PP,PD,allinfo,n):
    PPI = pasOrder.index(PP)
    PDI = pasOrder.index(PD)
    allOritime = 0
    allOpttime = 0
    allOridist = 0
    allOptdist = 0
    allOrifee = 0
    allOptfee = 0
    WT = allinfo['Walking Information'][PPI]["time"] + "," + allinfo['Walking Information'][PDI]["time"]
    WD = allinfo['Walking Information'][PPI]["distance"] + ","+ allinfo['Walking Information'][PDI]["distance"]
    for i in range(PPI+1,PDI+1):
        allOritime += int(allinfo['OriDrivingEachSection'][i]['time'])
        allOpttime += int(allinfo['OptDrivingEachSection'][i]['time'])
        allOridist += int(allinfo['OriDrivingEachSection'][i]['distance'])
        allOptdist += int(allinfo['OptDrivingEachSection'][i]['distance'])
        allOrifee += int(allinfo['OriDrivingEachSection'][i]['fee'])
        allOptfee += int(allinfo['OptDrivingEachSection'][i]['fee'])
        each = [{"Passenger%dTime"%(n):allOritime,"Passenger%dDist"%(n):allOridist,"Passenger%dFee"%(n):allOrifee,"Passenger%dWalkT"%(n):0,"Passenger%dWalkD"%(n):0},{"Passenger%dTime"%(n):allOpttime,"Passenger%dDist"%(n):allOptdist,"Passenger%dFee"%(n):allOptfee,"Passenger%dWTime"%(n):WT,"Passenger%dWDist"%(n):WD}]
    return each


def datainfo(dr,alt_dic):
    oporder = optimal_order_info(dr, alt_dic)[0]
    pasOrder = optimal_order_info(dr, alt_dic)[1]
    allinfo = all_info(oporder)
    P1 = each_pas_info(pasOrder,"P1_P","P1_D",allinfo,1)
    P2 = each_pas_info(pasOrder,"P2_P","P2_D",allinfo,2)
    P3 = each_pas_info(pasOrder,"P3_P","P3_D",allinfo,3)
    dicori = {**allinfo["OriginalRoute Information"], **P1[0],**P2[0],**P3[0]}
    dicopt = {**allinfo["OptimalRoute Information"][0], **P1[1],**P2[1],**P3[1]}
    dicori["Coordinate"] = ";".join(oporder)
    dicopt["Coordinate"] = ";".join(allinfo["OptimalRoute Information"][1])
    datainfo = [dicori,dicopt]
    return datainfo


def write_data_to_ExcleFile(alldatainfo,outPutFile):
    Lable1 = ['A','B','C','D','E','F',"G","H","I",'J','K', 'L','M',"N","O","P","Q","R","S","T","U","V","W"]
    wb = Workbook()
    sheet = wb.active
    sheet.title = "RoutingData"
    item_0 = alldatainfo[0]
    i = 0
    for key in item_0.keys():
        sheet[Lable1[i]+str(1)].value = key
        i = i+1
    j = 1
    for item in alldatainfo:
        k = 0
        for key in item:
            sheet[Lable1[k]+str(j+1)].value = item[key]
            k = k+1
        j = j+1
    wb.save(outPutFile)
    print("Excel Done!")


# ---------------------------------------------------  Main Part ---------------------------------------------------

all_dic = read_excel_to_dict("Thesis_LocationSet.xlsx")
print("Dictionary Completed！")
rd = random_dr_altdic(all_dic)
for i in [14,15,16]: #samples number
    rd = random_dr_altdic(all_dic)
    dr = rd[0]
    alt_dic = rd[1]
    alldatainfo = datainfo(dr,alt_dic)
    print("Number %d Information Completed!"%(i))
    write_data_to_ExcleFile(alldatainfo,'Routing%d.xlsx'%(i))
