import os
from os import environ
import sys
import winreg
from PyQt5 import uic
from PyQt5.Qt import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt
from winreg import *

import attachment
import connectDb
from PyQt5.QtWidgets import *
from PyQt5 import uic
from datetime import *
from datetime import datetime
from dateutil.relativedelta import relativedelta

import connect_ftp
import monthly_settlement
import display_tableWidget
import get_today
import recv_mail_daemon
import request_spend
import tot_profit_loss
import var_func
import update_version



def suppress_qt_warnings():
	environ["QT_DEVICE_PIXEL_RATIO"] = "0"
	environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
	environ["QT_SCREEN_SCALE_FACTORS"] = "1"
	environ["QT_SCALE_FACTOR"] = "1"

form_class = uic.loadUiType("mainwindow.ui")[0]


class MyWindow(QMainWindow, form_class):

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.go_update(0)
        self.go_login()

        self.brch_code.triggered.connect(self.go_branch_code)
        self.acc_code.triggered.connect(self.go_account_code)
        self.spcond_code.triggered.connect(self.go_spendcond_code)
        self.spend_code.triggered.connect(self.go_spend_code)
        self.crew_code.triggered.connect(self.go_crew_master)
        self.spend_cust_in.triggered.connect(self.go_spend_cust_in)
        self.r_spend.triggered.connect(self.go_request_spend)
        self.settlement_action.triggered.connect(self.go_settlement)
        self.actionUpdate.triggered.connect(self.go_update)
        self.tot_profit_loss.triggered.connect(self.go_totofprofitloss)
        self.config.triggered.connect(self.go_mailconf)

        dm = recv_mail_daemon.Checking_receivedMails()
        dm.daemon = True
        dm.start()

    def go_login(self):
        pop = login(self)
        pop.log_in()
        pop.exec_()
        self.permission = pop.permission

    def go_branch_code(self):
        permission = self.permission
        if permission == 'admin' or permission == 'manager' or permission == 'crew':
            pop = brach_Codemng(self)
            pop.begin(permission)
            pop.exec_()


    def go_account_code(self):
        permission = self.permission
        if permission == 'admin' or permission == 'manager' or permission == 'crew':
            pop = Account_codemng(self)
            pop.begin(permission,0)
            pop.exec_()


    def go_spendcond_code(self):
        permission = self.permission
        if permission == 'admin' or permission == 'manager' or permission == 'crew':
            pop = spend_cond_Mng(self)
            pop.begin(permission,"spend")
            pop.exec_()


    def go_crew_master(self):
        permission = self.permission
        if permission == 'admin' or permission == 'manager' or permission == 'crew':
            pop = crew_Mastermng(self)
            pop.begin(permission)
            pop.exec_()


    def go_spend_code(self):
        permission = self.permission
        pop = spend_Cust(self)
        pop.begin(permission)
        pop.exec_()

    def go_spend_cust_in(self):
        permission = self.permission
        pop = get_Cust_only(self)
        pop.begin(permission,0)
        pop.exec_()

    def go_request_spend(self):
        permission = self.permission
        pop = request_spend.request_Spend(self)
        pop.begin(permission)
        pop.exec_()

    def go_settlement(self):
        permission = self.permission
        pop = monthly_settlement.monthly_Settlement(self)
        pop.begin(permission)
        pop.exec_()

    def go_totofprofitloss(self):
        pop = tot_profit_loss.Profit_Loss(self)
        pop.begin()
        pop.exec_()

    def go_update(self,cond):
        # cond == 0 --> 자동업데이트 체크, cond == 1 --> 수동업데이트
        path = 'c:/samc/bsm/'
        executer = 'bsm_update.exe'
        sys_name = 'bsm'
        goup = update_version.update_Version()
        version = goup.get_version()
        self.setWindowTitle('BSM : ' + version)
        if cond == 0: #자동 업데이트 일때
            update = goup.check_update(sys_name)
            if update:
                goup.move_updatefile(path, sys_name, executer)
                goup.go_update(cond)
        else:
            goup.move_updatefile(path, sys_name, executer)
            update = goup.check_update(sys_name)
            if update:
                goup.go_update(cond)
            else:
                r = QMessageBox.question(self, "알림", "유효한 업데이트가 존재하지 않습니다. 그래도 업데이트 하시겠습니까?",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                if r == QMessageBox.Yes:
                    goup.go_update(cond)

    def go_mailconf(self):
        permission = self.permission
        if permission == 'admin':
            pop = mail_Conf(self)
            pop.begin(permission)
            pop.exec_()



######## 로그인 화면 ####################################################
class login(QDialog):
    def __init__(self, parent):
        super(login, self).__init__(parent)
        uic.loadUi("login.ui", self)
        self.show()

    def log_in(self):
        pass

    def log_inprocess(self):
        if self.checkBox.isChecked() == True:
            accs = self.get_logininfo()
            if accs:
                self.lineEdit.setText(accs[0])
                self.lineEdit_2.setText(accs[1])
        id = self.lineEdit.text()
        pwd = self.lineEdit_2.text()
        if id == '':
            QMessageBox.about(self, "알림", "계정정보 확인하세요")
            return
        conn = connectDb.connect_Db()
        curs = conn.cursor()
        sql = """select user_id,user_pwd,branch from user_master where user_id = '{}' and user_pwd = '{}'""".format(
            id, pwd)
        curs.execute(sql)
        res = curs.fetchone()
        if res:
            if res[2] == 'admin':
                self.permission = 'admin'
                self.close()
            elif res[2] == 'manager':
                self.permission = 'manager'
                self.close()
            elif res[2] == 'crew':
                self.permission = 'crew'
                self.close()
            else:
                self.permission = res[2]
                self.close()

    def get_logininfo(self):
        key = winreg.HKEY_CURRENT_USER
        key_value = "SOFTWARE\SAMC\local_inf\_b_m_s_info"
        try:
            reg = winreg.OpenKey(key, key_value, 0, winreg.KEY_ALL_ACCESS)
            user, type = winreg.QueryValueEx(reg, "user")
            pwd, type = winreg.QueryValueEx(reg, "pwd")
            cb, type = winreg.QueryValueEx(reg, "cb")

            if cb == '1':
                self.checkBox.setChecked(True)
                self.lineEdit.setText(user)
                self.lineEdit_2.setText(pwd)
            else:
                self.checkBox.setChecked(False)
        except FileNotFoundError:
            self.lineEdit.setText('')
            self.lineEdit_2.setText('')
            reg = winreg.CreateKey(key, key_value)
        winreg.CloseKey(reg)
        return [user,pwd]


##########매장코드 관리 ######################################################
class brach_Codemng(QDialog):
    def __init__(self, parent):
        super(brach_Codemng, self).__init__(parent)
        uic.loadUi("branch_code.ui", self)
        self.show()
        self.tableWidget.doubleClicked.connect(self.tbl_tblclicked)
        self.tableWidget.itemSelectionChanged.connect(self.tbl_itmchanged)
        self.tableWidget_attachment.doubleClicked.connect(self.tbl_attach_dblclicked)
        self.cat1.textChanged.connect(self.change_text_code)
        self.cat2.textChanged.connect(self.change_text_code)
        self.cat3.textChanged.connect(self.change_text_code)
        self.tableWidget_attachment.setContextMenuPolicy(Qt.ActionsContextMenu)
        del_action = QAction("삭제", self.tableWidget_attachment)
        del_action.triggered.connect(self.del_attachment)
        self.tableWidget_attachment.addAction(del_action)
        self.branch_cols = ['bcode', 'branch_name', 'contract_term','contract_term_end', 'lease','lease_end', 'running', 'franchise_mem','lease_mem', 'cat1', 'cat2', 'cat3']
        self.tsting = ','.join(self.branch_cols)
        self.edit = False
        self.tableWidget_attachment.setColumnWidth(0, 100)
        self.tableWidget_attachment.setColumnWidth(1, 600)

    def begin(self,permission):
        if permission == 'admin':
            self.pushButton_4.setEnabled(True)
            self.pushButton_5.setEnabled(True)
            self.pushButton_3.setEnabled(True)
        else:
            self.pushButton_4.setEnabled(False)
            self.pushButton_5.setEnabled(False)

        sql = self.make_sql('')
        self.sql = sql
        self.display_branch(sql)

    def make_sql(self,add_sql):
        tsting = self.tsting
        sql = """select {} from branch_master""".format(tsting)
        sql = sql + add_sql + """ order by bcode"""
        return sql

    def get_search(self):
        key_word = self.lineEdit.text()
        add_sql = """ where branch_name like '{}'""".format('%%' + key_word + '%%')
        sql = self.make_sql(add_sql)
        self.sql = sql
        self.display_branch(sql)

    def display_branch(self,sql):
        conn = connectDb.connect_Db()
        curs = conn.cursor()
        curs.execute(sql)
        res = curs.fetchall()
        self.res = res
        datas = []
        for i in range(len(res)):
            temp = [res[i][0],res[i][1],res[i][3],res[i][5],res[i][6],res[i][7]]
            datas.append(temp)
        tbl = self.tableWidget
        display_tableWidget.display(tbl,datas,'sql')
        if res:
            self.display_items(res[0])
            self.display_attachment()


    def display_items(self,res):
        cols = self.branch_cols
        for i in range(len(res)):
            if i == 6:
                getattr(self,cols[i]).setCurrentText(res[i])
            elif i == 0:
                continue
            elif i == 2 or i == 3 or i == 4 or i == 5:
                txt = self.reformat_text_date(res[i])
                if res[i] == '':
                    txt = ''
                getattr(self,cols[i]).setText(txt)
            else:
                getattr(self,cols[i]).setText(res[i])

    def reformat_text_date(self,text):
        txt = text[:4] + '-' + text[4:6] + '-' + text[6:]
        return txt


    def tbl_tblclicked(self):
        self.edit = True
        cols = self.branch_cols
        for i in range(len(cols)):
            if cols[i] == 'cat1' or cols[i] == 'cat2' or cols[i] == 'cat3':
                continue
            getattr(self, cols[i]).setEnabled(True)
        row = self.tableWidget.currentRow()
        res = self.res
        items = res[row]
        self.display_items(items)



    def tbl_itmchanged(self):
        row = self.tableWidget.currentRow()
        res = self.res
        bcode = self.tableWidget.item(row,0).text()
        for i in range(len(res)):
            if res[i][0] == bcode:
                items = res[i]
                break
        self.display_items(items)
        self.display_attachment()


    def add_save(self):
        if self.edit:
            cols = self.branch_cols
            tsting = self.tsting
            datas = []
            for i in range(len(cols)):
                if i == 6:
                    item = getattr(self,cols[i]).currentText()
                elif i == 7 or i == 8:
                    item = getattr(self,cols[i]).toPlainText()
                else:
                    item = getattr(self,cols[i]).text()
                    item = item.replace('/','')
                    item = item.replace('-','')
                getattr(self, cols[i]).setEnabled(False)
                datas.append(item)
            icode = self.bcode.text()
            sql1 = """delete from branch_master where bcode = '{}'""".format(icode)
            sql2 = """insert into branch_master ({}) values {}""".format(tsting,tuple(datas))
            conn = connectDb.connect_Db()
            curs = conn.cursor()
            curs.execute(sql1)
            curs.execute(sql2)
            conn.commit()
            conn.close()
            self.edit = False
            sql = self.sql
            self.display_branch(sql)
        else:
            self.edit = True
            cols = self.branch_cols
            for i in range(len(cols)):
                getattr(self, cols[i]).setEnabled(True)
        self.tbl_itmchanged()



    def change_text_code(self):
        code_text = self.cat1.text() + self.cat2.text() + self.cat3.text()
        self.bcode.setText(code_text)



    def display_attachment(self):
        row = self.tableWidget.currentRow()
        if row > -1:
            bcode = self.tableWidget.item(row,0).text()
        else:
            bcode = self.tableWidget.item(0,0).text()
        sql = """select category,filename,seq from attachment where branch_code = '{}'""".format(bcode)
        res = var_func.execute_sql(sql)
        self.att_res = res
        if res:
            tbl = self.tableWidget_attachment
            display_tableWidget.display(tbl, res, 'sql')
        else:
            self.tableWidget_attachment.setRowCount(0)


    def tbl_attach_dblclicked(self): # 더블 클릭하면 파일 오픈.....
        row = self.tableWidget_attachment.currentRow()
        filename  = self.tableWidget_attachment.item(row,1).text()
        dns = connect_ftp.Files_Manager()
        files = dns.get_files_of_ftp(filename)
        os.startfile((files))


    def attachment_mng(self): # 파일 첨부
        row = self.tableWidget.currentRow()
        if row > -1:
            bcode = self.tableWidget.item(row,0).text()
            filename = QFileDialog.getOpenFileNames(self, "파일선택", "c:\\","All kind of files (*.*)")
            if filename:
                filenames =filename[0]
                files = attachment.Files_attachment()
                files.begin(filenames,bcode,bcode,'매장정보')
                if files.complete:
                    QMessageBox.about(self,"저장성공","파일 첨부 완료 되었습니다.")
                    self.display_attachment()
        else:
            QMessageBox.about(self,"알림","선택된 지점이 없습니다.")
            return


    def del_attachment(self):
        res = self.att_res
        row = self.tableWidget_attachment.currentRow()
        seq = res[row][2]
        filename = self.tableWidget_attachment.item(row,1).text()
        files = connect_ftp.Files_Manager()
        res = files.get_delete_ftp(filename)
        if res:
            conn = connectDb.connect_Db()
            curs = conn.cursor()
            sql = """delete from attachment where seq = '{}'""".format(seq)
            curs.execute(sql)
            conn.commit()
            QMessageBox.about(self,"삭제","삭제되었습니다.")
            self.display_attachment()
        else:
            QMessageBox.about(self,"삭제싪패",'삭제할 파일이 해당 경로에 없습니다.')


    def quit(self):
        self.close()


################계정코드 관리 #########################################
class Account_codemng(QDialog):
    def __init__(self, parent):
        super(Account_codemng, self).__init__(parent)
        uic.loadUi("account.ui", self)
        self.show()
        self.acc_cols = ['acode','acc_name','cat1','cat2','cat3']
        self.tstring = ','.join(self.acc_cols)
        self.tableWidget.doubleClicked.connect(self.tbl_dblclicked)
        self.tableWidget.itemSelectionChanged.connect(self.tbl_itemchanged)
        self.cat1.textChanged.connect(self.make_code)
        self.cat2.textChanged.connect(self.make_code)
        self.cat3.textChanged.connect(self.make_code)
        self.edit = False
        self.tableWidget.setColumnWidth(0,70)
        self.tableWidget.setColumnWidth(1, 300)
        self.code = ''
        self.name = ''

    def begin(self,permission,cond):
        self.permission = permission
        if permission == 'admin':
            self.pushButton.setEnabled(True)
            self.pushButton.setEnabled(False)
        self.cond = cond
        self.permission = permission
        if permission == 'admin':
            self.pushButton.setEnabled(True)
        sql = self.sqls()
        self.display_account(sql)

    def sqls(self):
        sql = """select {} from account_master""".format(self.tstring)
        return sql

    def display_account(self,sql):
        conn = connectDb.connect_Db()
        curs = conn.cursor()
        curs.execute(sql)
        res = curs.fetchall()
        self.res = res
        tbl = self.tableWidget
        display_tableWidget.display(tbl,res,'sql')

    def tbl_dblclicked(self):
        permission = self.permission
        row = self.tableWidget.currentRow()
        cols = self.acc_cols
        res = self.res
        if self.cond == 1:
            self.code = self.tableWidget.item(row,0).text()
            self.name = self.tableWidget.item(row,1).text()
            self.close()
        for i in range(len(cols)):
            getattr(self,cols[i]).setText(res[row][i])
            if permission == 'admin':
                self.edit = True
                getattr(self,cols[i]).setEnabled(True)
        self.pushButton.setText('저장')
        self.pushButton.setEnabled(True)

    def tbl_itemchanged(self):
        row = self.tableWidget.currentRow()
        cols = self.acc_cols
        res = self.res
        for i in range(len(cols)):
            getattr(self, cols[i]).setText(res[row][i])


    def make_code(self):
        acc_code = self.cat1.text() + self.cat2.text() + self.cat3.text()
        self.acode.setText(acc_code)

    def add_item(self):
        cols = self.acc_cols
        tstring = self.tstring
        self.pushButton.setText('저장')
        if self.edit:
            acode = self.acode.text()
            datas = []
            for i in range(len(cols)):
                itm = getattr(self,cols[i]).text()
                datas.append(itm)
            sql1 = """delete from account_master where acode = '{}'""".format(acode)
            sql2 = """insert into account_master ({}) values {}""".format(tstring,tuple(datas))
            conn = connectDb.connect_Db()
            curs = conn.cursor()
            curs.execute(sql1)
            curs.execute(sql2)
            conn.commit()
            self.edit = False
            self.begin(self.permission,0)

        else:
            self.acode.setText('')
            self.acc_name.setText('')
            self.cat1.setText('')
            self.cat2.setText('')
            self.cat3.setText('')
            self.acode.setEnabled(True)
            self.acc_name.setEnabled(True)
            self.cat1.setEnabled(True)
            self.cat2.setEnabled(True)
            self.cat3.setEnabled(True)
            self.edit = True


    def quit(self):
        self.close()


#######지급방법, 지급조건 code ###############################################


class spend_cond_Mng(QDialog):
    def __init__(self, parent):
        super(spend_cond_Mng, self).__init__(parent)
        uic.loadUi("spcondmng.ui", self)
        self.show()
        self.tableWidget.doubleClicked.connect(self.tbl_dblclicked)
        self.tableWidget.itemSelectionChanged.connect(self.tbl_itemchanged)
        self.edit = False
        self.comboBox.currentIndexChanged.connect(self.change_idxcombo)
        self.bankcode = ''
        self.bankname = ''

    def begin(self,permission,cond): #cond : spend
        self.cond = cond
        self.permission = permission
        if permission == 'admin':
            self.pushButton.setEnabled(True)
            self.pushButton_2.setEnabled(True)
        if cond == 'payee':
            self.tableWidget.setHorizontalHeaderLabels(['은행코드', '은행이름'])
            self.comboBox.setCurrentIndex(2)
        else:
            self.tableWidget.setHorizontalHeaderLabels(['Code', '계정'])

        sql = self.sqls()
        self.display_spendpayment(sql)
        self.display_item(0)

    def sqls(self):
        idx = self.comboBox.currentIndex()
        if idx == 0:
            self.tbl_name = 'payment'
            sql = """select ccode,name from {}""".format(self.tbl_name)
        elif idx == 1:
            self.tbl_name = 'payment_order'
            sql = """select ccode,name from {}""".format(self.tbl_name)
        elif idx == 2:
            self.tableWidget.setHorizontalHeaderLabels(['은행코드', '은행이름'])
            self.tbl_name = 'banks'
            sql = """select bank_code,bank_name from {}""".format(self.tbl_name)

        return sql


    def display_spendpayment(self,sql):
        conn = connectDb.connect_Db()
        curs = conn.cursor()
        curs.execute(sql)
        res = curs.fetchall()
        self.res = res
        tbl = self.tableWidget
        display_tableWidget.display(tbl,res,'sql')


    def display_item(self,seq):
        self.ccode.setText('')
        self.name.setText('')
        res = self.res
        self.ccode.setText(res[seq][0])
        self.name.setText(res[seq][1])


    def tbl_dblclicked(self):
        permission = self.permission
        if permission == 'admin':
            self.ccode.setEnabled(True)
            self.name.setEnabled(True)
            self.comboBox.setEnabled(False)
            self.edit = True
        row = self.tableWidget.currentRow()
        if self.cond == 'payee':
            self.bankcode = self.tableWidget.item(row,0).text()
            self.bankname = self.tableWidget.item(row,1).text()
            self.close()
        self.display_item(row)

    def tbl_itemchanged(self):
        row = self.tableWidget.currentRow()
        self.display_item(row)


    def add_save(self):
        if self.edit:
            ccode = self.ccode.text()
            name = self.name.text()
            sql1 = """delete from {} where ccode = '{}'""".format(self.tbl_name,ccode)
            sql2 = """insert into {} (ccode, name) values {}""".format(self.tbl_name,(ccode,name))
            conn = connectDb.connect_Db()
            curs = conn.cursor()
            curs.execute(sql1)
            curs.execute(sql2)
            conn.commit()
            sql = """select ccode,name from {}""".format(self.tbl_name)
            self.comboBox.setEnabled(True)
            self.ccode.setEnabled(True)
            self.name.setEnabled(True)
            self.edit = False
            self.display_spendpayment(sql)
        else:
            self.edit = True
            self.ccode.setEnabled(True)
            self.name.setEnabled(True)
            self.ccode.setText('')
            self.name.setText('')

    def change_idxcombo(self):
        permission = self.permission
        self.begin(permission,self.cond)


    def quit(self):
        self.close()

##########직원관리 폼 crew_master ##########################################
class crew_Mastermng(QDialog):
    def __init__(self, parent):
        super(crew_Mastermng, self).__init__(parent)
        uic.loadUi("crew_master.ui", self)
        self.show()
        self.crew_cols = ['branch_name', 'crew_name', 'crew_contact', 'inprocess', 'service_period', 'service_gbn',
                          'account_no', 'account_bank', 'payday', 'service', 'crew_address', 'remark', 'outprocess',
                          'branch_code','crew_code','mails','bank_code']
        self.tstring = ','.join(self.crew_cols)
        self.branches = self.get_branchnames()
        self.tableWidget.doubleClicked.connect(self.tbl_dblclicked)
        self.tableWidget.itemSelectionChanged.connect(self.tbl_itemchanged)
        self.service.currentIndexChanged.connect(self.change_service_cond)
        self.branch_name.currentIndexChanged.connect(self.change_branch)
        self.edit = False
        self.update = False
        self.outprocess.setDate(QDate(0,0,0))
        self.inprocess.dateChanged.connect(self.change_workingday)
        self.tableWidget_attachment.doubleClicked.connect(self.tbl_attach_dblclicked)
        self.tableWidget_attachment.setContextMenuPolicy(Qt.ActionsContextMenu)
        del_action = QAction("삭제", self.tableWidget_attachment)
        del_action.triggered.connect(self.del_attachment)
        self.tableWidget_attachment.addAction(del_action)
        self.tableWidget_attachment.setColumnWidth(0, 100)
        self.tableWidget_attachment.setColumnWidth(1, 300)
        self.tableWidget_attachment.setColumnWidth(2, 100)
        self.att_res = False

    def get_branchnames(self):
        conn = connectDb.connect_Db()
        curs = conn.cursor()
        sql = """select branch_name,bcode from branch_master"""
        curs.execute(sql)
        res = curs.fetchall()
        temp = []
        branches = []
        self.branch_name.addItem('지점선택')
        for i in range(len(res)):
            if res[i][0] not in temp:
                self.branch_name.addItem(res[i][0])
                temp.append(res[i][0])
                temp1 = [res[i][0], res[i][1]]
                branches.append(temp1)
        return branches


    def begin(self,permission):
        self.permission = permission
        if permission == 'admin':
            self.pushButton_3.setEnabled(True)
        sql = self.sqls('')
        res = self.get_sqlresult(sql)
        self.add_branch_tocombo(res)
        self.display_crews(res)


    def add_branch_tocombo(self,res):
        self.comboBox.clear()
        temp = []
        self.comboBox.addItem('전체')
        for i in range(len(res)):
            branch = res[i][0]
            if branch not in temp:
                temp.append(branch)
                self.comboBox.addItem(branch)


    def sqls(self,add_sql):
        tstring = self.tstring
        sql = """select {} from crew_master""".format(tstring)
        sql = sql + add_sql
        return sql


    def get_sqlresult(self,sql):
        conn = connectDb.connect_Db()
        curs = conn.cursor()
        curs.execute(sql)
        res = curs.fetchall()
        conn.close()
        return res


    def display_crews(self,res):
        self.res = res
        tbl = self.tableWidget
        self.display(tbl,res,'sql')


    def display(self,tbl, res, data):  # res 는 튜플 또는 리스트 data 는 data  타입
        tbl.setRowCount(len(res))
        tbl.verticalHeader().setVisible(False)
        if data == 'sql':  # res는 튜플
            for i in range(len(res)):
                for ii in range(len(res[i])):
                    if res[i][ii] == 'NULL' or res[i][ii] == '' or res[i][ii] is None:
                        tbl.setItem(i, ii, QTableWidgetItem(''))
                    else:
                        if ii == 5 and res[i][ii] == '0':
                            texts = '정직'
                        elif ii == 5 and res[i][ii] == '1':
                            texts = '계약'
                        else:
                            texts = str(res[i][ii])
                            texts = texts.replace('\\n', ' ')
                        tbl.setItem(i, ii, QTableWidgetItem(texts))

    def get_search(self):
        crewname = self.lineEdit.text()
        branch = self.comboBox.currentText()
        if branch == '전체':
            add_sql = """ where crew_name like '{}'""".format('%%' + crewname + '%%')
        else:
            add_sql = """ where branch_name = '{}' and crew_name like '{}'""".format(branch,'%%' + crewname + '%%')
            if crewname == '':
                add_sql = """ where branch_name = '{}'""".format(branch)
        sql = self.sqls(add_sql)
        res = self.get_sqlresult(sql)
        self.display_crews(res)

    def display_items(self,row,permission):
        res = self.res
        if res:
            cols = self.crew_cols
            for i in range(len(cols)):
                if cols[i] == 'service_gbn':
                    if res[row][i] == '0': #service_gbn 1 -- >정직 2-->알바
                        getattr(self,cols[i]).setCurrentText('정직')
                    else:
                        getattr(self, cols[i]).setCurrentText('계약')
                elif cols[i] == 'service' or cols[i] == 'branch_name':
                    getattr(self,cols[i]).setCurrentText(res[row][i])
                elif cols[i] == 'inprocess':
                    tdate = res[row][3]
                    if tdate == '':
                        continue
                    self.inprocess.setDate(QDate(int(tdate[:4]), int(tdate[4:6]), int(tdate[6:])))
                elif cols[i] == 'outprocess':
                    tdate = res[row][12]
                    if tdate == '':
                        continue
                    self.outprocess.setDate(QDate(int(tdate[:4]), int(tdate[4:6]), int(tdate[6:])))
                else:
                    getattr(self,cols[i]).setText(res[row][i])
                if permission == 'admin':
                    if cols[i] != 'outprocess':
                        if not self.edit:
                            getattr(self,cols[i]).setEnabled(False)
                        else:
                            getattr(self,cols[i]).setEnabled(True)

    def get_banks(self):
        pop = spend_cond_Mng(self)
        pop.begin(self.permission,cond = 'payee')
        pop.exec_()
        bank_cd = pop.bankcode
        bank = pop.bankname
        self.account_bank.setText(bank)
        self.bank_code.setText(bank_cd)

    def tbl_dblclicked(self):
        permission = self.permission
        if permission == 'admin':
            self.edit = True
        row = self.tableWidget.currentRow()
        self.display_items(row,permission)
        self.pushButton_3.setEnabled(True)
        self.pushButton_4.setEnabled(True)
        self.pushButton_5.setEnabled(True)
        self.tableWidget_attachment.setEnabled(True)
        self.display_items(row,permission)

    def tbl_itemchanged(self):
        row = self.tableWidget.currentRow()
        self.display_items(row, False)
        self.display_attachment()


    def change_service_cond(self):
        cond = self.service.currentText()
        if cond == 'N':
            self.outprocess.setEnabled(True)
        else:
            self.outprocess.setEnabled(False)

    def add_save(self):
        cols = self.crew_cols
        permission = self.permission
        if permission == 'admin' or 'manager':
            if self.edit:
                branch_name = self.branch_name.currentText()
                if branch_name == '지점선택':
                    QMessageBox.about(self, "알림", "지점선택이 되어 있지 않습니다. ")
                save_confirm = self.save_confirmdata()
                if not save_confirm:
                    QMessageBox.about(self, "확인", "신규 직원 저장시 신분증,사업자등록증, 통장사본중 하나는 꼭 첨부되어야 합니다.")
                    datas = []
                for i in range(len(cols)):
                    if cols[i] == 'service_gbn':
                        idx = getattr(self,cols[i]).currentIndex()
                        item = str(idx)
                    elif cols[i] == 'service' or cols[i] == 'branch_name':
                        item = getattr(self,cols[i]).currentText()
                    elif cols[i] == 'remark':
                        item = getattr(self,cols[i]).toPlainText()
                    elif cols[i] == 'branch_code' :
                        item = getattr(self,cols[i]).text()
                    elif cols[i] == 'inprocess':
                        tdate1 = self.inprocess.date()
                        item = tdate1.toString('yyyyMMdd')
                    elif cols[i] == 'outprocess':
                        if self.service.currentText() == 'N':
                            item = ''
                        else:
                            tdate1 = self.outprocess.date()
                            item = tdate1.toString('yyyyMMdd')
                    elif cols[i] == 'service_period':
                        if self.service.currentText() == 'Y':
                            item = getattr(self, cols[i]).text()
                        else:
                            item = ''
                    else:
                        item = getattr(self,cols[i]).text()
                    getattr(self,cols[i]).setEnabled(False)
                    self.edit = False
                    datas.append(item)
                self.save_data(datas)
                self.save_info_attachment()
                self.make_payee(datas)
                self.pushButton_4.setEnabled(False)
                self.pushButton_5.setEnabled(False)
                self.tableWidget_attachment.setEnabled(True)
                self.get_search()
                #else:
                    #pass
                    #QMessageBox.about(self, "확인","신규 직원 저장시 신분증,사업자등록증, 통장사본중 하나는 꼭 첨부되어야 합니다.")
            else:
                self.edit = True
                ncode = self.make_crewcode()
                self.pushButton_4.setEnabled(True)
                self.pushButton_5.setEnabled(True)
                self.tableWidget_attachment.setEnabled(True)
                for i in range(len(cols)):
                    if cols[i] == 'crew_code':
                        getattr(self,cols[i]).setText(ncode)
                    elif cols[i] == 'branch_code':
                        continue
                    elif cols[i] not in  ['outprocess','crew_code']:
                        getattr(self,cols[i]).setEnabled(True)
                    elif cols[i] == 'outprocess':
                        self.outprocess.setDate(QDate(1900,1,1))
                    elif cols[i] == 'inprocess':
                        tdate = get_today.get_date()
                        self.inprocess.setDate(QDate(int(tdate[:4]), int(tdate[4:6]), int(tdate[6:])))
                    elif cols[i] == 'branch_name':
                        branches = self.branches
                        self.branch_code.setText(branches[0][1])
                    if cols[i] not in ['outprocess','inprocess','crew_code','service','service_gbn','branch_name']:
                        getattr(self,cols[i]).setText('')

#['branch_name', 'crew_name', 'crew_contact', 'inprocess', 'service_period', 'service_gbn',
 #                         'account_no', 'account_bank', 'payday', 'service', 'crew_address', 'remark', 'outprocess',
 #                         'branch_code','crew_code','mails','bank_code']

    def save_data(self,datas):

        #self.crew_cols = ['branch_name', 'crew_name', 'crew_contact', 'inprocess', 'service_period', 'service_gbn',
        #                  'account_no', 'account_bank', 'payday', 'service', 'crew_address', 'remark', 'outprocess',
        #                  'branch_code', 'crew_code']

        conn = connectDb.connect_Db()
        curs = conn.cursor()
        tstring = self.tstring
        branch_code = datas[13]
        branch_name = datas[0]
        crew_name = datas[1]
        crew_contact = datas[2]
        crew_code = datas[14]
        mails = datas[15]
        account_no = datas[6]
        account_bank = datas[7]
        bank_code = datas[16]
        porder_code = '15'
        if datas[5] == '0':
            porder_name = '정규직급여/소사장배당'
            account_code = 'CH3001'
            account_name = '정규직급여'
        else:
            account_code = 'CH3002'
            account_name = '계약직급여'
            porder_name = '계약직급여'
        #if datas[9] != 'Y':
            #self.update_payee_noworking(crew_code)
        pcode = '20'
        pname = '인터넷뱅킹'

        payee_datas = ('H' + crew_code, crew_name,crew_name,crew_contact,mails,branch_code,branch_name,account_no,account_bank,bank_code,porder_code,porder_name,pcode,pname,account_code,account_name)
        sql_payee1 = """delete from payee_cust where pcode = '{}'""".format('H' + crew_code)
        sql_payee2 = """insert into payee_cust (pcode,pname,pceo,tel,mails,branch_code,branch_name,account_no,account_bank,bank_code,payment_order_code,payment_order_name,payment_code, payment_name,account_code,account_name) values {}""".format(payee_datas)
        curs.execute(sql_payee1)
        curs.execute(sql_payee2)

        sql1 = """delete from crew_master where crew_code = '{}'""".format(crew_code)
        sql2 = """insert into crew_master ({}) values {}""".format(tstring,tuple(datas))
        curs.execute(sql1)
        curs.execute(sql2)
        conn.commit()
        conn.close()

    #def update_payee_noworking(self,crew_code):
    #    sql = """update payee set confirm = 'N' where gcode = '{}'""".format(crew_code)
    #    var_func.execute_sql_insdelupd(sql)

    def save_confirmdata(self):
        count_attachment = self.tableWidget_attachment.rowCount()
        tags = []
        for row in range(count_attachment):
            txt = self.tableWidget_attachment.cellWidget(row,2).currentText()
            tags.append(txt)
        if '신분증' in tags or '사업자등록증' in tags or '통장사본' in tags:
            return True
        else:
            return False



    def save_info_attachment(self):
        att_res = self.att_res
        if att_res:
            tbl = self.tableWidget_attachment
            rows = tbl.rowCount()
            if rows > 0:
                for row in range(rows):
                    tag = tbl.cellWidget(row,2).currentText()
                    sql = """update attachment set tag = '{}' where seq = '{}'""".format(tag,att_res[row][3])
                    conn = connectDb.connect_Db()
                    curs = conn.cursor()
                    curs.execute(sql)
                conn.commit()
                conn.close()

##### 지급처 MAster 자동 생성
    def make_payee(self,datas):


        #self.paycust_cols = ['gcode', 'gname', 'account_no', 'account_bank', 'branch_code', 'branch_name',
                             # 'account_code','account_name','payment_order_code', 'payment_order_name','payment_name','confirm','remark','pay','mngcode','payment_code','bank_code']


        mngcode = self.payee_get_mngno()
        gcode = 'H' + str(datas[14])
        gname = datas[1]
        account_no = str(datas[6])
        account_bank = str(datas[7])
        branch_code = str(datas[13])
        branch_name = str(datas[0])
        service = datas[9]

        if datas[5] == '1':
            account_code = 'CH3002'
            account_name = '계약직급여'
            payment_order_name = '계약직급여'
        else:
            account_code = 'CH3001'
            account_name = '정직원급여'
            payment_order_name = '정규직급여/소사장배당'

        payment_order_code = '15'
        payment_name = '인터넷뱅킹'
        if service == 'Y':
            confirm = 'Y'
        else:
            confirm = 'N'
        remark = ''
        pay = ''
        payment_code = '20'
        bank_code = str(datas[16])
        datas = [gcode, gname, account_no, account_bank, branch_code, branch_name, account_code, account_name,
                 payment_order_code, payment_order_name, payment_name, confirm, remark, pay, mngcode,
                 payment_code, bank_code]

        self.save_payee_datas_sql(datas,gcode)

    def payee_get_mngno(self):
        sql = """select max(mngcode) from payee"""
        res = var_func.execute_sql(sql)
        if res[0][0]:
            no = int(res[0][0]) + 1
        else:
            no = 1
        no = str(no).zfill(6)
        return no

    def save_payee_datas_sql(self,datas,gcode):
        payee_cols = ['gcode', 'gname', 'account_no', 'account_bank', 'branch_code', 'branch_name',
                             'account_code', 'account_name', 'payment_order_code', 'payment_order_name', 'payment_name',
                             'confirm', 'remark', 'pay', 'mngcode', 'payment_code', 'bank_code']
        tstring = ','.join(payee_cols)
        sql1 = """delete from payee where gcode = '{}'""".format(gcode)
        sql2 = """insert into payee ({}) values {}""".format(tstring,tuple(datas))
        conn = connectDb.connect_Db()
        curs = conn.cursor()
        curs.execute(sql1)
        curs.execute(sql2)
        conn.commit()
        conn.close()




    def make_crewcode(self):
        conn = connectDb.connect_Db()
        curs = conn.cursor()
        sql = """SELECT  max(a.crew_code) from crew_master a"""
        curs.execute(sql)
        res = curs.fetchone()
        if res[0] is not None:
            last_code = int(res[0])
            new_code = str(last_code + 1).zfill(4)
        else:
            new_code = '0001'
        return new_code

    
    def change_workingday(self):
        today = datetime.today()
        working = self.service.currentText()
        if working == 'Y':
            tdate = self.inprocess.date()
            tdate = tdate.toPyDate()
            working_day = str(today.date() - tdate)
            self.service_period.setText(working_day)
        else:
            tdate = self.outprocess.date()
            if tdate != '':
                tdate = tdate.toPyDate()
                working_day = str(today.date() - tdate)
                self.service_period.setText(working_day)


    def change_branch(self):
        branches = self.branches
        text = self.branch_name.currentText()
        for i in range(len(branches)):
            if text == branches[i][0]:
                self.branch_code.setText(branches[i][1])
                break

    def display_attachment(self):
        bcode = self.branch_code.text()
        hcode = self.crew_code.text()
        sql = """select category,filename,tag,seq from attachment where branch_code = '{}' and category = '직원' and key_code = '{}'""".format(bcode,hcode)
        res = var_func.execute_sql(sql)
        self.att_res = res
        if res:
            tbl = self.tableWidget_attachment
            display_tableWidget.display(tbl, res, 'sql')
        else:
            self.tableWidget_attachment.setRowCount(0)
        rows = self.tableWidget_attachment.rowCount()
        for row in range(rows):
            self.cb = QComboBox()
            self.cb.addItem('신분증')
            self.cb.addItem('사업자등록증')
            self.cb.addItem('통장사본')
            self.cb.addItem('기타')
            items = tbl.item(row,2).text()
            if items is None or items == '':
                self.cb.setCurrentText('')
            else:
                self.cb.setCurrentText(items)
            self.tableWidget_attachment.setCellWidget(row, 2, self.cb)


    def tbl_attach_dblclicked(self): # 더블 클릭하면 파일 오픈.....
        row = self.tableWidget_attachment.currentRow()
        filename  = self.tableWidget_attachment.item(row,1).text()
        dns = connect_ftp.Files_Manager()
        files = dns.get_files_of_ftp(filename)
        os.startfile((files))


    def attachment_mng(self): # 파일 첨부
        row = self.tableWidget.currentRow()
        if row > -1:
            bcode = self.branch_code.text()
            hcode = self.crew_code.text()
            filename = QFileDialog.getOpenFileNames(self, "파일선택", "c:\\","All kind of files (*.*)")
            if filename:
                filenames =filename[0]
                files = attachment.Files_attachment()
                files.begin(filenames,bcode,hcode,'직원')
                if files.complete:
                    QMessageBox.about(self,"저장성공","파일 첨부 완료 되었습니다.")
                    self.display_attachment()
        else:
            QMessageBox.about(self,"알림","선택된 지점이 없습니다.")
            return


    def del_attachment(self):
        permission = self.permission
        if permission == 'admin':
            res = self.att_res
            row = self.tableWidget_attachment.currentRow()
            seq = res[row][3]
            filename = self.tableWidget_attachment.item(row,1).text()
            files = connect_ftp.Files_Manager()
            res = files.get_delete_ftp(filename)
            if res:
                conn = connectDb.connect_Db()
                curs = conn.cursor()
                sql = """delete from attachment where seq = '{}'""".format(seq)
                curs.execute(sql)
                conn.commit()
                QMessageBox.about(self,"삭제","삭제되었습니다.")
                self.display_attachment()
            else:
                QMessageBox.about(self,"삭제실패","삭제할 파일이 없습니다.")



    def quit(self):
        self.close()

################## 지급처 CODE 관리 #########################
class spend_Cust(QDialog):
    def __init__(self, parent):
        super(spend_Cust, self).__init__(parent)
        uic.loadUi("spend_cust.ui", self)
        self.show()
        self.paycust_cols = ['gcode', 'gname', 'account_no', 'account_bank', 'branch_code', 'branch_name',
                              'account_code','account_name','payment_order_code', 'payment_order_name','payment_name','confirm','remark','pay','mngcode','payment_code','bank_code']
        self.tstring = ','.join(self.paycust_cols)
        self.tableWidget.doubleClicked.connect(self.tbl_dblclicked)
        self.tableWidget.itemSelectionChanged.connect(self.item_changed)
        self.edit = False
        self.make_paymentorder_combo()
        self.make_payment_combo()
        self.payment_order_name.currentIndexChanged.connect(self.payment_order_changed)
        self.payment_name.currentIndexChanged.connect(self.payment_changed)
        #self.payment_order_code.currentIndexChanged.connect(self.payment_order_code_changed)
        self.seq = ''



    def begin(self,permission):
        self.permission = permission
        if permission  == 'admin':
            self.pushButton.setEnabled(True)
            self.pushButton_5.setEnabled(True)
            self.pushButton_6.setEnabled(True)
            self.pushButton_7.setEnabled(True)
        add_sql = ''
        self.sqls(add_sql,permission)
        self.display_grid()
        self.display_items(0)


    def sqls(self, add_sql,permission):
        tstring = self.tstring
        if permission != 'admin' and permission != 'manager' and permission != 'crew':
            sql = """select {} from payee where branch_code = '{}'""".format(tstring,permission)
        else:
            sql = """select {} from payee""".format(tstring)
        sql = sql + add_sql
        self.sql = sql


    def display_grid(self):
        sql = self.sql
        conn = connectDb.connect_Db()
        curs = conn.cursor()
        curs.execute(sql)
        res = curs.fetchall()
        self.res = res
        if res:
            tbl = self.tableWidget
            display_tableWidget.display(tbl,res,'sql')


    def display_items(self,row):
        if self.res:
            items = self.res[row]
            self.seq = self.res[row][13]
            cols = self.paycust_cols
            for i in range(len(items)):
                if cols[i] == 'payment_order_name' or cols[i] == 'payment_name' or cols[i] == 'payment_order_code':
                    getattr(self,cols[i]).setCurrentText(items[i])
                    #self.payment_order_changed()
                elif cols[i] == 'confirm':
                    continue
                else:
                    getattr(self,cols[i]).setText(items[i])


    def get_search(self): # 조회
        permission = self.permission
        keyword = self.search_edit.text()
        if keyword != '':
            idx = self.comboBox.currentIndex()
            if idx == 0:
                add_sql = """ where branch_name like '{}'""".format('%%' + keyword + '%%')
            elif idx == 1:
                add_sql = """ where gname like '{}'""".format('%%' + keyword + '%%')
            elif idx == 2:
                add_sql = """ where account_name like '{}'""".format('%%' + keyword + '%%')
        else:
            add_sql = ''
        if permission != 'admin' and permission != 'manager':
            add_sql = ""
        self.sqls(add_sql,permission)
        self.display_grid()
        self.display_items(0)


    def go_findcust(self):
        pop = get_Cust(self)
        pop.begin(self.permission)
        pop.exec_()
        if pop.bcode:
            bcode = pop.bcode
            branch_name = pop.branch_name
            self.branch_code.setText(bcode)
            self.branch_name.setText(branch_name)
            return [bcode,branch_name]

    def go_account(self):
        pop = Account_codemng(self)
        pop.begin(False,1)
        pop.exec_()
        self.account_code.setText(pop.code)
        self.account_name.setText(pop.name)


    def get_bank(self):
        pop = spend_cond_Mng(self)
        pop.begin(self.permission,"payee")
        pop.exec_()
        self.account_bank.setText(pop.bankname)
        self.bank_code.setText(pop.bankcode)


    def add_save(self):
        permission = self.permission
        cols = self.paycust_cols
        tstring = self.tstring
        if self.edit:
            pcode = self.gcode.text()
            sql = """select pcode from payee_cust where pcode = '{}'""".format(pcode)
            res = self.execute_sql(sql)
            if res:
                datas = []
                for i in range(len(cols)):
                    if cols[i] == 'payment_order_name' or cols[i] == 'payment_name' or cols[i] == 'payment_order_code':
                        item = getattr(self, cols[i]).currentText()
                        getattr(self, cols[i]).setEnabled(False)
                    elif cols[i] == 'remark':
                        item = getattr(self, cols[i]).toPlainText()
                        getattr(self, cols[i]).setEnabled(False)
                    elif cols[i] == 'confirm':
                        if permission == 'admin':
                            item = 'Y'
                        else:
                            item = 'N'
                    else:
                        item = getattr(self, cols[i]).text()
                        getattr(self,cols[i]).setEnabled(False)
                    datas.append(item)
                self.save_datas_sql(datas)
                self.display_grid()
            else:
                QMessageBox.about(self,"경고","등록되지 않은 매장명입니다. 매장을 확인하세요")
            self.edit = False
            self.pushButton_6.setEnabled(False)
            self.pushButton_7.setEnabled(False)
        else:
            mngno = self.get_mngno()
            self.mngcode.setText(mngno)
            branches = self.go_findcust()
            for i in range(len(cols)):
                if cols[i] == 'payment_order_name' or cols[i] == 'payment_name' or cols[i] == 'payment_order_code':
                    getattr(self,cols[i]).setCurrentText('')
                elif cols[i] == 'mngcode' or cols[i] == 'confirm':
                    continue
                else:
                    getattr(self,cols[i]).setText('')
                getattr(self,cols[i]).setEnabled(True)
            self.edit = True
            self.branch_code.setText(branches[0])
            self.branch_name.setText(branches[1])
            self.pushButton_6.setEnabled(True)
            self.pushButton_7.setEnabled(True)


    def save_datas_sql(self,datas):
        tstring = self.tstring
        mngno = self.mngcode.text()
        sql1 = """delete from payee where mngcode = '{}'""".format(mngno)
        sql2 = """insert into payee ({}) values {}""".format(tstring,tuple(datas))
        conn = connectDb.connect_Db()
        curs = conn.cursor()
        curs.execute(sql1)
        curs.execute(sql2)
        conn.commit()
        conn.close()

    def get_mngno(self):
        sql = """select max(mngcode) from payee"""
        res = self.execute_sql(sql)
        if res[0][0]:
            no = int(res[0][0]) + 1
        else:
            no = 1
        no = str(no).zfill(6)
        return no


    def tbl_dblclicked(self): #수정모드
        permission = self.permission
        row = self.tableWidget.currentRow()
        self.pushButton_6.setEnabled(True)
        self.pushButton_7.setEnabled(True)
        if row > -1:
            if permission == 'admin':
                self.edit = True
                cols = self.paycust_cols
                for i in range(len(cols)):
                    if cols[i] in ['seq','confirm','bank_code']:
                        continue
                    getattr(self,cols[i]).setEnabled(True)
        self.display_items(row)

    def item_changed(self):
        row = self.tableWidget.currentRow()
        res = self.res
        self.display_items(row)
        #items = res[row]
        #for i in range(len(items[0])):
        #    if items[i] == 'payment_order_name' or items[i] == 'payment_name':
        #        getattr(self,items[i]).setCurrentText(items[i])
        #    else:
        #        getattr(self,items[i]).setText(items[i])

    def make_paymentorder_combo(self):
        sql = """select ccode, name from payment_order"""
        pay_orders = self.execute_sql(sql)
        self.pay_orders = pay_orders
        temp = []
        for x in pay_orders:
            if x[0] not in temp:
                temp.append(x[0])
                self.payment_order_name.addItem(x[1])

    def make_payment_combo(self):
        sql = """select ccode, name from payment"""
        payment = self.execute_sql(sql)
        self.payment = payment
        temp = []
        for x in payment:
            if x[0] not in temp:
                temp.append(x[0])
                self.payment_name.addItem(x[1])
        self.payment_name.setCurrentIndex(0)
        self.payment_code.setText(self.payment_name.currentText())

    def execute_sql(self,sql):
        conn = connectDb.connect_Db()
        curs = conn.cursor()
        curs.execute(sql)
        res = curs.fetchall()
        conn.close()
        return res

    def payment_order_changed(self):
        idx = self.payment_order_name.currentIndex()
        pay_orders = self.pay_orders
        text = pay_orders[idx][0]
        self.payment_order_code.setCurrentText(text)


    def payment_changed(self):
        idx = self.payment_name.currentIndex()
        payment = self.payment
        text = payment[idx][0]
        self.payment_code.setText(text)



    def go_payee_cust(self):
        permission = self.permission
        cond = 0 # 화면우측 상단에서 누르면 단순 등록 모드로 접근 ---> 구분 0
        pop = get_Cust_only(self)
        pop.begin(permission,cond)
        pop.exec_()

    def go_payee_cust2(self):
        permission = self.permission
        cond = 1 # 화면 좌측 하단에서 누르면 단순 등록 모드로 접근 ---> 구분 0
        pop = get_Cust_only(self)
        pop.begin(permission,cond)
        pop.exec_()
        if pop.pcode:
            self.pcode = pop.r_pcode
            self.pname = pop.r_pname
            self.acc_no = pop.r_account_no
            self.acc_bank = pop.r_account_bank
            self.acc_code = pop.r_account_code
            self.acc_name = pop.r_account_name
            self.pay_order_name = pop.r_payment_order_name
            self.pay_order_code = pop.r_payment_order_code
            self.pay_name = pop.r_payment_name
            self.pay_code = pop.r_payment_code
            self.pay_amt = pop.r_pay
            self.bnk_code = pop.r_bank_code

            self.gcode.setText(self.pcode)
            self.gname.setText(self.pname)
            self.account_no.setText(self.acc_no)
            self.account_name.setText(self.acc_name)
            self.account_bank.setText(self.acc_bank)
            self.account_code.setText(self.acc_code)
            self.payment_order_code.setCurrentText(self.pay_order_code)
            self.payment_order_name.setCurrentText(self.pay_order_name)
            self.payment_code.setText(self.pay_code)
            self.payment_name.setCurrentText(self.pay_name)
            self.pay.setText(self.pay_amt)
            self.bank_code.setText(self.bnk_code)


    def quit(self):
        self.close()


## 지급처 코드에서 지점 가져오기 선택

class get_Cust(QDialog):
    def __init__(self, parent):
        super(get_Cust, self).__init__(parent)
        uic.loadUi("spend_cust_popup_get_branch.ui", self)
        self.show()
        self.tableWidget.doubleClicked.connect(self.tbl_dblclicked)
        self.bcode = False
        self.branch_name = False

    def begin(self,permission):
        self.permission = permission
        if permission != 'admin' and permission != 'manager':
            sql = """select bcode, branch_name from branch_master where bcode = '{}'""".format(permission)
        else:
            sql = """select bcode, branch_name from branch_master"""
        res = self.execute_sql(sql)
        tbl = self.tableWidget
        display_tableWidget.display(tbl,res,'sql')

    def get_search(self):
        permission = self.permission
        if permission != 'admin' and permission != 'manager':
            return
        keyword = self.lineEdit.text()
        sql = """select bcode, branch_name from branch_master where branch_name like '{}'""".format('%%' + keyword + '%%')
        res = self.execute_sql(sql)
        tbl = self.tableWidget
        display_tableWidget.display(tbl, res, 'sql')


    def tbl_dblclicked(self):
        row = self.tableWidget.currentRow()
        bcode = self.tableWidget.item(row,0).text()
        branch_name = self.tableWidget.item(row,1).text()
        self.bcode = bcode
        self.branch_name = branch_name
        self.close()


    def execute_sql(self, sql):
        conn = connectDb.connect_Db()
        curs = conn.cursor()
        curs.execute(sql)
        res = curs.fetchall()
        return res

    def quit(self):
        self.bcode = False
        self.close()



###### 지급처만 등록 #################
class get_Cust_only(QDialog):
    def __init__(self, parent):
        super(get_Cust_only, self).__init__(parent)
        uic.loadUi("payee_cust.ui", self)
        self.show()
        self.cols = ['pcode', 'pname', 'pbizno', 'pceo', 'tel', 'fax', 'mails', 'remark', 'account_bank', 'account_no',
                     'account_code', 'account_name', 'payment_order_name', 'payment_order_code', 'payment_name',
                     'payment_code','pay','bank_code']
        self.tstring = ','.join(self.cols)
        self.tableWidget.doubleClicked.connect(self.tbl_dblclicked)
        self.tableWidget.itemSelectionChanged.connect(self.tbl_itemchanged)
        self.make_combo_pordername()
        self.payment_order_name.currentIndexChanged.connect(self.porder_changed)
        self.payment_name.currentIndexChanged.connect(self.p_changed)
        self.make_combo_pname()
        self.edit = False

        self.r_pcode = ''
        self.r_pname = ''
        self.r_account_bank = ''
        self.r_account_no = ''
        self.r_account_code = ''
        self.r_account_name = ''
        self.r_payment_order_name = ''
        self.r_payment_order_code = ''
        self.r_payment_name = ''
        self.r_payment_code = ''
        self.r_pay = ''
        self.r_bank_code = ''

    def make_combo_pordername(self):
        sql = """select ccode,name from payment_order"""
        res = self.execute_sql(sql)
        self.po_res = res
        temp = []
        for i in range(len(res)):
            if res[i][0] not in temp:
                temp.append(res[i][0])
                self.payment_order_name.addItem(res[i][1])


    def make_combo_pname(self):
        sql = """select ccode,name from payment"""
        res = self.execute_sql(sql)
        self.pres = res
        temp = []
        for i in range(len(res)):
            if res[i][0] not in temp:
                temp.append(res[i][0])
                self.payment_name.addItem(res[i][1])
        self.payment_name.setCurrentIndex(0)
        self.payment_code.setText(self.payment_name.currentText())

    def porder_changed(self):
        res = self.po_res
        text = self.payment_order_name.currentText()
        for i in range(len(res)):
            if text == res[i][1]:
                self.payment_order_code.setText(res[i][0])

    def p_changed(self):
        res = self.pres
        text = self.payment_name.currentText()
        for i in range(len(res)):
            if text == res[i][1]:
                self.payment_code.setText(res[i][0])



    def begin(self,permission,cond):
        if permission == 'admin':
            self.pushButton_4.setEnabled(True)
            self.pushButton_5.setEnabled(True)
            self.pushButton_2.setEnabled(True)
        self.cond = cond
        self.permission = permission
        if permission == 'admin':
            self.pushButton_2.setEnabled(True)
            self.permission = permission
        add_sql = ''
        self.sqls(add_sql)
        self.display_grid()
        self.display_item(0)

    def sqls(self,add_sql):
        tstring = self.tstring
        sql = """select {} from payee_cust""".format(tstring)
        sql = sql + add_sql
        self.sql = sql + """ order by pcode"""

    def display_grid(self):
        sql = self.sql
        res = self.execute_sql(sql)
        self.res = res
        tbl = self.tableWidget
        display_tableWidget.display(tbl,res,'sql')


    def display_item(self,row):
        res = self.res
        cols = self.cols
        if res:
            items = res[row]
            for i in range(len(items)):
                if cols[i] == 'payment_order_name' or cols[i] == 'payment_name':
                    getattr(self,cols[i]).setCurrentText(items[i])
                else:
                    getattr(self,cols[i]).setText(items[i])

    def get_search(self):
        idx = self.comboBox.currentIndex()
        word = self.lineEdit.text()
        sql_where = " where "
        if idx == 0:
            add_sql = sql_where + """pname like '{}'""".format('%%' + word + '%%')
        elif idx == 1:
            add_sql = sql_where + """pceo = '{}'""".format(word)
        elif idx == 2:
            add_sql = sql_where + """pbizno = '{}'""".format(word)
        elif idx == 3:
            add_sql = sql_where + """tel = '{}'""".format(word)
        self.sqls(add_sql)
        self.display_grid()
        self.display_item(0)

    def tbl_dblclicked(self):
        self.edit = True
        row = self.tableWidget.currentRow()
        self.display_item(row)
        permission = self.permission
        cond = self.cond
        if permission == 'admin':
            cols = self.cols
            for i in range(len(cols)):
                getattr(self,cols[i]).setEnabled(True)
        if cond == 1: #더블클릭시 값 가져가야 하는 경우
            self.r_pcode = self.pcode.text()
            self.r_pname = self.pname.text()
            self.r_account_bank = self.account_bank.text()
            self.r_account_no = self.account_no.text()
            self.r_account_code = self.account_code.text()
            self.r_account_name = self.account_name.text()
            self.r_payment_order_name = self.payment_order_name.currentText()
            self.r_payment_order_code = self.payment_order_code.text()
            self.r_payment_name = self.payment_name.currentText()
            self.r_payment_code = self.payment_code.text()
            self.r_pay = self.pay.text()
            self.r_bank_code = self.bank_code.text()

            self.close()



    def tbl_itemchanged(self):
        row = self.tableWidget.currentRow()
        self.display_item(row)

    def add_save(self):
        cols = self.cols
        datas = []
        permission = self.permission
        if permission == 'admin':
            if self.edit:
                for i in range(len(cols)):
                    getattr(self,cols[i]).setEnabled(False)
                    if cols[i] == 'remark':
                        item = getattr(self,cols[i]).toPlainText()
                    elif cols[i] == 'payment_order_name' or cols[i] == 'payment_name':
                        item = getattr(self,cols[i]).currentText()
                    else:
                        item = getattr(self,cols[i]).text()
                    datas.append(item)
                self.save_datas(datas)
                self.display_grid()
                self.edit = False
            else:
                for i in range(len(cols)):
                    getattr(self,cols[i]).setEnabled(True)
                    if cols[i] == 'payment_order_name' or cols[i] == 'payment_name':
                        getattr(self,cols[i]).setCurrentIndex(0)
                    elif cols[i] == 'pcode':
                        new_code = self.get_cust_code()
                        self.pcode.setText(new_code)
                        self.pcode.setEnabled(False)
                    else:
                        getattr(self,cols[i]).setText('')
                    self.edit = True


    def save_datas(self,datas):
        tstring = self.tstring
        pcode = self.pcode.text()
        sql1 = """delete from payee_cust where pcode = '{}'""".format(pcode)
        sql2 = """insert into payee_cust ({}) values {}""".format(tstring, tuple(datas))
        conn = connectDb.connect_Db()
        curs = conn.cursor()
        curs.execute(sql1)
        curs.execute(sql2)
        conn.commit()
        conn.close()



    def get_cust_code(self):
        sql = """select max(pcode) from payee_cust where pcode like '1%%'"""
        res = var_func.execute_sql(sql)
        maxval = int(res[0][0])
        newcode = maxval + 1
        return str(newcode)



    def execute_sql(self, sql):
        conn = connectDb.connect_Db()
        curs = conn.cursor()
        curs.execute(sql)
        res = curs.fetchall()
        conn.close()
        return res

    def get_account(self):
        pop = Account_codemng(self)
        pop.begin(False,1)
        pop.exec_()
        self.code = pop.code
        self.name = pop.name
        self.account_code.setText(self.code)
        self.account_name.setText(self.name)

    def get_banks(self):
        pop = spend_cond_Mng(self)
        pop.begin(self.permission,"payee")
        pop.exec_()
        self.account_bank.setText(pop.bankname)
        self.bank_code.setText(pop.bankcode)

    def quit(self):
        self.pcode = False
        self.close()



class mail_Conf(QDialog):
    def __init__(self, parent):
        super(mail_Conf, self).__init__(parent)
        uic.loadUi("mail_conf.ui", self)
        self.show()


    def begin(self,permission):
        if permission != 'admin':
            return
        sql = """select ceo_mail,financial_mail,financial_cc_mail from mail_conf"""
        res = var_func.execute_sql(sql)
        self.lineEdit.setText(res[0][0])
        self.lineEdit_2.setText(res[0][1])
        self.lineEdit_3.setText(res[0][2])


    def save(self):
        ceo_mail = self.lineEdit.text()
        financial_mail = self.lineEdit_2.text()
        financial_cc_mail = self.lineEdit_3.text()
        sql = """update mail_conf set ceo_mail = '{}', financial_mail = '{}', financial_cc_mail = '{}' where seq = 1""".format(ceo_mail,financial_mail,financial_cc_mail)
        try:
            var_func.execute_sql_insdelupd(sql)
            QMessageBox.about(self,"Save Complete","저장됬다.")
            self.close()
        except:
            QMessageBox.about(self, "Save Fail", "저장안됬다.")



    def quit(self):
        self.close()


if __name__ == "__main__":
    suppress_qt_warnings()
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
