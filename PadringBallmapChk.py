# encoding = utf-8
__author__ = 'Cyan'

import sys
import random
import string
import xlrd
import re

dbg_on = 1

class ReadInputFile():
    '''
    read input excel
    get table data
    get specified area data and return it as list
    3 kind of specified area:
        power pin
        padring(arealist and emptylist)
        ballmap
    '''
    global dbg_on
    def __init__(self, fpath, fname):
        self.fpath = fpath
        self.fname = fname
        self.fullpath = '%s%s'%(self.fpath, self.fname)

    def open_excel(self, path):
        try:
            data = xlrd.open_workbook(path)
            return data
        except Exception as e:
            print(e)

    def get_area_data(self, dtype, sheet, pos):
        self.dtype = dtype
        self.sheet = sheet
        self.pos = pos

        area_details={'namelist':[], 'loclist':[], 'padnamelist':[], 'padnumlist':[]}
        data = self.open_excel(self.fullpath)
        table = data.sheet_by_name(self.sheet)
        area_data = []
        '''
        if dbg_on > 0:
            print(self.pos)'''
        for i in range(self.pos['rowstart'],(self.pos['rowstart']+self.pos['rowdelta'])):
            area_data.append(table.row_values(i))
        '''
        if dbg_on > 0:
            print(area_data)
            print('000000000000000')'''
        area_details = self.proc_area_data(self.dtype, area_data)
        '''
        if dbg_on > 0:
            print(area_details)
            print('111111111111111')'''
        return area_details

    #power pin name detect
    def ballname_det(self, strname):
        '''
        if dbg_on > 0:
            print(strname)'''

        #ballname should have more than 2 bits
        if len(strname)<2:
            return 0
        #ballname should start with A-Z
        if re.match(r'[^A-Z]', strname[0]):
            return 0

        for i in strname:
            '''
            if dbg_on > 0:
                print(i)'''
            i_not_str = re.match(r'\W', i)
            i_is_lowercase = re.match(r'[a-z]', i)
            #if any string not in A-Z or 0-9 or _:
            if i_is_lowercase or i_not_str:
                return 0

        #ballname should end with A-Z or 0-9
        if strname[-1] == '_':
            return 0

        return 1

    #location name detect
    def locname_det(self, strname):
        if strname == '':
            return 0
        '''
        if dbg_on > 0:
            print(strname)'''
        str_without_space=''
        for i in strname:
            if re.match(r'\w',i) or i==',':
                str_without_space = '%s%s'%(str_without_space, i)
        '''
        if dbg_on > 0:
            print(str_without_space)'''
        tmpstrlist = str_without_space.split(',')
        '''
        if dbg_on > 0:
            print(tmpstrlist)'''
        for i in tmpstrlist:
            '''
            if dbg_on > 0:
                print(i)
            '''
            j_list=[]
            for j in i:
                j_list.append(j)
            if len(j_list)<2 or len(j_list)>3:
                return 0
            j_list_0_not_str = re.match(r'[^A-Z]', j_list[0])
            if j_list_0_not_str:
                return 0

            if len(j_list) == 2:
                j_list_1_illegal = re.match(r'[^1-9]', j_list[1])
                j_list_2_illegal = ''
            else:#elif len(j_list) == 3:
                j_list_1_illegal = re.match(r'[^1]', j_list[1])
                j_list_2_illegal = re.match(r'[^0-9]', j_list[2])

            if j_list_1_illegal or j_list_2_illegal:
                return 0

        return 1


    def proc_area_data(self, dtype, datalist):
        if dtype == 'powerpin':
            area_details={'namelist':[], 'loclist':[], 'padnamelist':[], 'padnumlist':[]}
            NameColFlag = 1
            '''
            if dbg_on > 0:
                print(len(datalist))'''
            for i in range(len(datalist)):

                #for j in datalist[i]:
                '''
                if dbg_on > 0:
                    print(len(datalist[i]))'''
                for j in range(len(datalist[i])):
                    '''
                    if dbg_on > 0:
                        print(datalist[i][j])'''
                    if NameColFlag == 1:#waiting for name
                        if self.ballname_det(datalist[i][j])>0:
                            NameColFlag+=1
                            area_details['namelist'].append(datalist[i][j])
                    elif NameColFlag == 2:#waiting for location
                        if self.locname_det(datalist[i][j])>0:
                            tmpstr = datalist[i][j]

                            str_without_space=''
                            for k in tmpstr:
                                if re.match(r'\w',k) or k==',':
                                    str_without_space = '%s%s'%(str_without_space, k)

                            tmpstrlist = str_without_space.split(',')
                            if tmpstrlist:
                                area_details['loclist'].append(tmpstrlist)
                            else:
                                area_details['loclist'].append([])
                                print('Warning! Location cell is empty!!! i is %d, j is %d'%(i, j))
                            NameColFlag -= 1

            '''
            if dbg_on > 0:
                print(area_details['namelist'])
                print(area_details['loclist'])'''

        elif dtype == 'ballmap':
            area_details={'namelist':[], 'loclist':[], 'padnamelist':[], 'padnumlist':[]}
            j_list = []
            row_loc = ''
            col_loc = ''
            curr_loc = ''
            for j in range(len(datalist[0])):
                j_list.append(datalist[0][j])
            for i in range(1, len(datalist)):
                for j in range(len(datalist[i])):
                    if self.ballname_det(datalist[i][j]):
                        area_details['namelist'].append(datalist[i][j])
                        row_loc=datalist[i][0]
                        col_loc=str(int(j_list[j]))
                        curr_loc='%s%s'%(row_loc, col_loc)

                        if dbg_on > 0:
                            print(curr_loc)
                        area_details['loclist'].append(curr_loc)

        elif dtype == 'padring':
            pad_num_sequence_error=0
            area_details={'namelist':[], 'loclist':[], 'padnamelist':[], 'padnumlist':[]}
            j_list=[]
            emptylist=[]

            ball_num_col=0
            pad_num_col=1
            pad_name_col=2
            ball_name_col=7
            for j in range(len(datalist[0])):
                j_list.append(datalist[0][j])
            ##ball_num_col = j_list.index('')
            #ball_name_col = j_list.index('Ball Name')
            #pad_num_col = j_list.index('Pad #')
            #pad_name_col = j_list.index('Pad Name on Chip')
            for i in range(1, len(datalist)):
                if i != int(datalist[i][pad_num_col]):
                    pad_num_sequence_error+=1

                if self.locname_det(datalist[i][ball_num_col])>0:
                    area_details['namelist'].append(datalist[i][ball_name_col])
                    area_details['loclist'].append(datalist[i][ball_num_col])
                    area_details['padnamelist'].append(datalist[i][pad_name_col])
                    area_details['padnumlist'].append(int(datalist[i][pad_num_col]))
                elif self.ballname_det(datalist[i][ball_num_col])>0 and datalist[i][ball_num_col] == datalist[i][ball_name_col]:
                    area_details['namelist'].append(datalist[i][ball_name_col])
                    area_details['loclist'].append(datalist[i][ball_num_col])
                    area_details['padnamelist'].append(datalist[i][pad_name_col])
                    area_details['padnumlist'].append(int(datalist[i][pad_num_col]))
                else:
                    '''
                    if dbg_on>0:
                        print(datalist[i])'''
                    emptylist.append(datalist[i])

            if pad_num_sequence_error>0:
                print('WARNING!!! pad_num_sequence_error! error num is %d'%pad_num_sequence_error)

            print('>>>>>>>>>>>>>>>> BEGIN: PRINT EMPTY PAD NOT CONNECTED TO BALL <<<<<<<<<<<<<<<<')
            for i in emptylist:
                print(i)
            print('>>>>>>>>>>>>>>>>  END: PRINT EMPTY PAD NOT CONNECTED TO BALL  <<<<<<<<<<<<<<<<')

        return area_details


def pp_to_bm_chk(pplist, bmlist):
    ErrCode=0

    if len(pplist['namelist']) != len(pplist['loclist']):
        ErrCode=1000000
        print('ERROR - 1 -1! power pin namelist mismatch with loclist! .')

    ErrCodePre=ErrCode

    for i in range(len(pplist['namelist'])):
        search_num = bmlist['namelist'].count(pplist['namelist'][i])
        if search_num==0:
            print('ERROR - 1 - 2! power pin name not exist in ballmap!')
            ErrCode+=1
        elif search_num==1:
            search_loc=bmlist['namelist'].index(pplist['namelist'][i])
            if len(pplist['loclist'][i])>1:
                print('ERROR - 2 - 1! power pin name exist in ballmap, but power pin location is too much more than ballmap!')
                ErrCode+=10
            elif pplist['loclist'][i][0] != bmlist['loclist'][search_loc]:
                print('ERROR - 2 - 2! power pin name exist in ballmap, but power pin location is different with ballmap!')
                ErrCode+=100
        else:#elif search_num>1:
            if  len(pplist['loclist'][i]) > search_num:
                print('ERROR - 2 - 1A! power pin name exist in ballmap, but power pin location is too much more than ballmap!')
                ErrCode+=1000
            elif  len(pplist['loclist'][i]) < search_num:
                print('ERROR - 2 - 3! power pin name exist in ballmap, but power pin location is less than ballmap!')
                ErrCode+=10000
            else:
                for j in range(len(bmlist['namelist'])):
                    if pplist['namelist'][i] == bmlist['namelist'][j]:
                        if bmlist['loclist'][j] not in pplist['loclist'][i]:
                            print('ERROR - 2 - 2A! power pin name exist in ballmap, but power pin location is different with ballmap!')
                            ErrCode+=100000

        if ErrCodePre != ErrCode:
            ErrCodePre = ErrCode
            print('ppname is %s'%(pplist['namelist'][i]))

    if ErrCode > 0:
        return('Powerpin to Ballmap check has ERROR, error code is %07d'%(ErrCode))
    else:
        return 'Powerpin to Ballmap check is OK'

def pr_to_bm_chk(prlist, bmlist, pplist):
    pr_pp_list = {'namelist':[], 'loclist':[], 'padnumlist':[], 'multicnt':[]}
    ErrCode = 0
    ErrCodePre = ErrCode

    '''
    first foreach the prlist,
    kick off power pins to pr_pp_list,
    check the rest pads's name and location
    '''
    for i in range(len(prlist['namelist'])):
        if prlist['namelist'][i] == prlist['loclist'][i]:
            if pr_pp_list['namelist'].count(prlist['namelist'][i]) == 0:
                pr_pp_list['namelist'].append(prlist['namelist'][i])
                pr_pp_list['loclist'].append(prlist['loclist'][i])
                pr_pp_list['padnumlist'].append(prlist['padnumlist'][i])
                pr_pp_list['multicnt'].append(1)
            else:
                inx_tmp = pr_pp_list['namelist'].index(prlist['namelist'][i])
                pr_pp_list['multicnt'][inx_tmp]+=1
        else:
            search_num = bmlist['namelist'].count(prlist['namelist'][i])
            if search_num == 0:
                print('ERROR - 1! pad not exist in ballmap!')
                ErrCode +=1
            elif search_num == 1:
                search_loc = bmlist['namelist'].index(prlist['namelist'][i])
                if prlist['loclist'][i] != bmlist['loclist'][search_loc]:
                    print('ERROR - 2! pad exist in ballmap, but the location is different!')
                    ErrCode +=10
            else:#elif search_name > 1, normally this is a power pin name.
                print('ERROR - 3! pad is not a pp, but it has multiple balls connecting!')
                ErrCode +=100

        if ErrCodePre != ErrCode:
            ErrCodePre = ErrCode
            print('pad num is %s'%(prlist['padnumlist'][i]))
            
    '''
    if dbg_on>0:
        print('>>>>>>>>>>>>>>>> BEGIN: PRINT POWER PIN LIST FROM PARING <<<<<<<<<<<<<<<<')
        print(pr_pp_list)
        print('>>>>>>>>>>>>>>>>  END: PRINT POWER PIN LIST FROM PARING  <<<<<<<<<<<<<<<<')
    '''

    '''
    then foreach the pr_pp_list,
    check it with pplist and bmlist
    '''
    for i in range(len(pr_pp_list['namelist'])):
        num_in_pp = pplist['namelist'].count(pr_pp_list['namelist'][i])
        num_in_bm = bmlist['namelist'].count(pr_pp_list['namelist'][i])

        if num_in_bm == 0:
            print('ERROR - 1! pad(pp) not exist in ballmap!')
            ErrCode +=1000
        else:
            if num_in_pp == 0:
                print('ERROR - 4! pad(pp) exist in ballmap, but not exist in pplist!')
                ErrCode +=1000
            elif num_in_pp > 1:
                print('ERROR - 5! pad(pp) exist in ballmap, but has mutiple exist in pplist!')
                ErrCode +=10000
            else:#elif num_in_pp==1:
                index_in_pp =  pplist['namelist'].index(pr_pp_list['namelist'][i])
                if num_in_bm > len(pplist['loclist'][index_in_pp]):
                    print('ERROR - 6-1! pad(pp) exist in ballmap, but its location description in pplist is less than ballmap!')
                    print()
                    ErrCode +=10000
                elif num_in_bm < len(pplist['loclist'][index_in_pp]):
                    print('ERROR - 6-2! pad(pp) exist in ballmap, but its location description in pplist is too much more than ballmap!')
                    ErrCode +=100000
                else:
                    for j in pplist['loclist'][index_in_pp]:
                        if j not in bmlist['loclist']:
                            print('ERROR - 7! pad(pp) exist in ballmap, but its location description in pplist is not exist in ballmap!')
                            ErrCode +=1000000
                        else:
                            inx_tmp = bmlist['loclist'].index(j)
                            if pplist['namelist'][index_in_pp] != bmlist['namelist'][inx_tmp]:
                                print('ERROR - 8! pad(pp) exist in ballmap, but its location description in pplist is not correct in ballmap!')
                                ErrCode +=10000000

        if ErrCodePre != ErrCode:
            ErrCodePre = ErrCode
            #inx_tmp = prlist['namelist'].index(pr_pp_list['namelist'][i])
            #print('pad num is %s'%(prlist['padnumlist'][inx_tmp]))
            print('pad num is %d'%(pr_pp_list['padnumlist'][i]))

    if ErrCode>0:
        return('Padring to Ballmap check has ERROR, error code is %08d'%(ErrCode))
    else:
        return 'Padring to Ballmap check is OK'

def bm_to_pr_chk(bmlist, prlist, pplist):
    ErrCode = 0
    ErrCodePre = ErrCode
    bm_pp_list = {'namelist':[], 'loclist':[], 'multicnt':[]}
    '''
    first compare bmlist with pplist
    then kick pp out
    then compare the other balls with prlist
    '''
    for i in range(len(bmlist['namelist'])):
        num_in_pp = pplist['namelist'].count(bmlist['namelist'][i])
        num_in_pr = prlist['namelist'].count(bmlist['namelist'][i])
        if num_in_pr == 0:
            if num_in_pp == 0:
                print('ERROR - 1! ballname not exist in padring!')
                ErrCode+=1
            elif num_in_pp > 1:
                print('ERROR - 1 - 1! ballname not exist in padring but has mutiple descriptions in powerpin area!')
                ErrCode+=100
            else:
                if bm_pp_list['namelist'].count(bmlist['namelist'][i]) == 0:
                    bm_pp_list['namelist'].append(bmlist['namelist'][i])
                    bm_pp_list['loclist'].append([bmlist['loclist'][i]])
                else:#count==1
                    inx_tmp = bm_pp_list['namelist'].index(bmlist['namelist'][i])
                    bm_pp_list['loclist'][inx_tmp].append(bmlist['loclist'][i])
        else:
            if num_in_pp > 1:
                print('ERROR - 2! ballname exist in padring but has mutiple descriptions in powerpin area!')
                ErrCode+=1000
            elif num_in_pp == 1:
                if bm_pp_list['namelist'].count(bmlist['namelist'][i]) == 0:
                    bm_pp_list['namelist'].append(bmlist['namelist'][i])
                    bm_pp_list['loclist'].append([bmlist['loclist'][i]])
                else:#count==1
                    inx_tmp = bm_pp_list['namelist'].index(bmlist['namelist'][i])
                    bm_pp_list['loclist'][inx_tmp].append(bmlist['loclist'][i])
            else:#not pp
                if num_in_pr > 1:
                    print('ERROR - 3! ballname not PP but has mutiple match in padring!')
                    ErrCode+=10000
                else:#num_in_pr == 1
                    inx_tmp = prlist['namelist'].index(bmlist['namelist'][i])
                    if bmlist['loclist'][i] != prlist['loclist'][inx_tmp]:
                        print('ERROR - 4! ballname exist in padring but location not match!')
                        ErrCode+=100000

        if ErrCodePre != ErrCode:
            ErrCodePre = ErrCode
            print('current ball name is %s, ball loc is %s'%(bmlist['namelist'][i], bmlist['loclist'][i]))


    if dbg_on>0:
        print('>>>>>>>>>>>>>>>> BEGIN: PRINT POWER PIN LIST FROM BALLMAP <<<<<<<<<<<<<<<<')
        print(bm_pp_list)
        print('>>>>>>>>>>>>>>>>  END: PRINT POWER PIN LIST FROM BALLMAP  <<<<<<<<<<<<<<<<')


    #these balls are confirmed existed 1 time in pplist at above logic.
    for i in range(len(bm_pp_list['namelist'])):
        idx_tmp = pplist['namelist'].index(bm_pp_list['namelist'][i])
        if len(bm_pp_list['loclist'][i]) < len(pplist['loclist'][idx_tmp]):
            print('ERROR - 5-A! ballname exist in PP but location num is less than PP descrption!')
            ErrCode+=1000000
        elif len(bm_pp_list['loclist'][i]) > len(pplist['loclist'][idx_tmp]):
            print('ERROR - 5-B! ballname exist in PP but location num is more than PP descrption!')
            ErrCode+=10000000
        else:
            for j in bm_pp_list['loclist'][i]:
                if j not in pplist['loclist'][idx_tmp]:
                    print('ERROR - 6! ballname exist in PP but location is different with PP descrption!')
                    ErrCode+=100000000

        if ErrCodePre != ErrCode:
            ErrCodePre = ErrCode
            print('current ball name is %s, ball loc is %s, pp loc is %s,'%(bm_pp_list['namelist'][i],bm_pp_list['loclist'][i],pplist['loclist'][idx_tmp]))

    if ErrCode>0:
        return('Ballmap to Padring check has ERROR, error code is %09d'%(ErrCode))
    else:
        return 'Ballmap to Padring check is OK'


def printtest():
    '''
    read input excel file
    get specified area data
    '''
    print('############################################################')
    print('......Start the padring excel auto check flow!')
    print('############################################################')

    #RIF = ReadInputFile('D:\\py_cases\\case6_padring_ballmap_autochk\\', 'Nile_padring_floorplan_ballmap_1.4.xls')
    #pp_area = RIF.get_area_data('powerpin', 'Padring', {'rowstart':17, 'rowdelta':21})
    #bm_area = RIF.get_area_data('ballmap', 'Ball Map', {'rowstart':8, 'rowdelta':9})
    #pr_area = RIF.get_area_data('padring', 'Padring', {'rowstart':54, 'rowdelta':307})

    sys.stdout.write('Please input padring_ballmap file name(format:Nile_padring_floorplan_ballmap_1.4.xls):')
    tmp_str = input()
    RIF = ReadInputFile('', tmp_str)


    (get_input, pp_rs, pp_re) = (1, 1, 1)
    while(get_input):
        sys.stdout.write('Please input powerpin area\'s start row number and end row number(format:aaa,bbb):')
        str_tmp = input()
        if str_tmp != '':
            str_tmplist = str_tmp.split(',')
            if len(str_tmplist) == 2:# and int(str_tmplist[0]) < int(str_tmplist[1]):
                pp_rs =  int(str_tmplist[0])
                pp_re = int(str_tmplist[1])
                if str(pp_rs) == str_tmplist[0] and str(pp_re) == str_tmplist[1]:
                    if pp_re > pp_rs:
                        get_input = 0

        if get_input > 0:
            print('input error, please retry.')

    pp_rs_real = pp_rs - 1
    pp_rd_real = pp_re + 1 - pp_rs


    pp_area = RIF.get_area_data('powerpin', 'Padring', {'rowstart':pp_rs_real, 'rowdelta':pp_rd_real})

    (get_input, pp_rs, pp_re) = (1, 1, 1)
    while(get_input):
        sys.stdout.write('Please input ballmap area\'s start row number and end row number(format:aaa,bbb):')
        str_tmp = input()
        if str_tmp != '':
            str_tmplist = str_tmp.split(',')
            if len(str_tmplist) == 2:# and int(str_tmplist[0]) < int(str_tmplist[1]):
                pp_rs =  int(str_tmplist[0])
                pp_re = int(str_tmplist[1])
                if str(pp_rs) == str_tmplist[0] and str(pp_re) == str_tmplist[1]:
                    if pp_re > pp_rs:
                        get_input = 0

        if get_input > 0:
            print('input error, please retry.')

    pp_rs_real = pp_rs - 1
    pp_rd_real = pp_re + 1 - pp_rs - 1

    bm_area = RIF.get_area_data('ballmap', 'Ball Map', {'rowstart':pp_rs_real, 'rowdelta':pp_rd_real})

    (get_input, pp_rs, pp_re) = (1, 1, 1)
    while(get_input):
        sys.stdout.write('Please input padring area\'s start row number and end row number(format:aaa,bbb):')
        str_tmp = input()
        if str_tmp != '':
            str_tmplist = str_tmp.split(',')
            if len(str_tmplist) == 2:# and int(str_tmplist[0]) < int(str_tmplist[1]):
                pp_rs =  int(str_tmplist[0])
                pp_re = int(str_tmplist[1])
                if str(pp_rs) == str_tmplist[0] and str(pp_re) == str_tmplist[1]:
                    if pp_re > pp_rs:
                        get_input = 0

        if get_input > 0:
            print('input error, please retry.')

    pp_rs_real = pp_rs - 1
    pp_rd_real = pp_re + 1 - pp_rs

    pr_area = RIF.get_area_data('padring', 'Padring', {'rowstart':pp_rs_real, 'rowdelta':pp_rd_real})
    print('')

    if dbg_on>0:
        print(pp_area)
        print(bm_area)
        print(pr_area)


    '''
    do check
    STEP1, powerpin to ballmap check
    check items:
        name match;
        location match;
    '''
    chk_rslt=''
    chk_rslt=pp_to_bm_chk(pp_area, bm_area)
    print('############################################################')
    print(chk_rslt)
    print('############################################################')
    print('')

    '''
    do check
    STEP2, padring to ballmap check
    check items:
        (include a padring to power pin check)
        name match;
            if name has location, then location match;
            if name has powerpin, redirect to the powerpin location, then location match;
    '''
    chk_rslt=''
    chk_rslt=pr_to_bm_chk(pr_area, bm_area, pp_area)
    print('############################################################')
    print(chk_rslt)
    print('############################################################')
    print('')

    '''
    do check
    STEP3, ballmap to padring check
    check items:
        (include a ballmap to power pin check)
        if name not match pp,
            name match with padring;
            location match with padring;
        if name match with pp,
            location also match with pp.
    '''
    chk_rslt=''
    chk_rslt=bm_to_pr_chk(bm_area, pr_area, pp_area)
    print('############################################################')
    print(chk_rslt)
    print('############################################################')
    print('')

def main():
    printtest()

if __name__ == '__main__':
    main()
