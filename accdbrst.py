# -*- coding: utf-8 -*-

"""
Designed Based on pyodbc package.
if user is not pursuing speed of code or handling big data, then this module could help user 
manipulate access more easily and briefly

Modified log:
Created by Henry Gao on the 11th Jan 2019
Modified by Henry Gao on the 22th Jan 2019: add function(getfieldname)
Modified by Henry Gao on the 25th Jan 2019: modify function(importfromexcel),allow user set imported table name and decide Header.
Modified by Henry Gao on the 25th Jan 2019: add function(droptable)
Modified by Henry Gao on the 28th Jan 2019: add function(importfromcsv)
Modified by Henry Gao on the 4th Mar 2019: add function(importfromtxt)
"""
import pyodbc

class Access():

    def __init__(self,path):
        
        self.path=path

        self.conn=pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+self.path+';')

        self.cursor=self.conn.cursor()
        
    def openrecordset(self,sql):
        """
        return: recordset
        type: accdb recordset
        
        """
        cursor=self.cursor

        cursor.execute(sql)
        
        recordset=cursor.fetchall()

        return recordset
        
    def execute(self,*sqls):
        """
        run sqls
        it can handle multi sqls in a row
        and it could be faster if write code as following:
        sql1=''
        sql2=''
        sql3-''
        .execute(sql1,sq12,sql1)
        howver this style is not good for error handling since 
        i will raise error on the same line
        """
        cursor=self.cursor
        
        for sql in list(sqls):
             
            cursor.execute(sql)

        cursor.commit()

    def executemany(self,sql,dataset):

        cursor=self.cursor
        
        cursor.executemany(sql,dataset)
    
        cursor.commit()

    def importfromexcel(self,wbpath,shtname,temp_name,header=False):
        """
        default setting hearder is False, if source file includes header
        user need to set header=True
        """
        cursor = self.cursor

        try:

            cursor.execute("Drop Table "+temp_name+"")

        except:

            pass

        if header==False:
            sql=r"SELECT * INTO "+temp_name+" FROM [Excel 12.0 Xml;HDR=NO;IMEX=2;ACCDB=YES;DATABASE="+wbpath+"].["+shtname+"$]"
        else:
            sql=r"SELECT * INTO "+temp_name+" FROM [Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE="+wbpath+"].["+shtname+"$]"      
        
        cursor.execute(sql)
        
        cursor.commit()
            
    def closecurrentdb(self):
        
        conn=self.conn

        cursor=self.cursor

        cursor.close()

        conn.close()

    def tablelist(self):

        conn=self.conn

        cursor=self.cursor

        namelist=[]

        for rcd in list(cursor.tables()):

            name=rcd[2]

            namelist.append(name)

        return namelist
        
    def getfieldname(self,tablename):
        """
        return field name of a given table
        type:list
        """
        conn=self.conn

        cursor=self.cursor

        res = cursor.execute("SELECT * FROM "+tablename+"")

        fieldlist = [tuple[0] for tuple in res.description]

        return fieldlist

    def droptable(self,*tables):
        """
        drop multiple tables
        style refers to run sqls
        """
        cursor=self.cursor

        for table in list(tables):

            try:
                cursor.execute("Drop Table "+table+"")
            except:
                raise Exception("Unable to drop "+table+". Maybe currently in use or not exsited")
    
    def importfromcsv(self,wbpath,table_name):

        import csv

        cursor = self.cursor

        with open(wbpath, 'r') as csvfile:
            
            rd = csv.reader(csvfile)
            
            row=next(rd)
            
            query = 'insert into '+table_name+' values ({0})'
            #use sql: insert into table values(?,?,?) style
            query = query.format(','.join('?' * len(row)))

            for row in rd:
                try:
                    cursor.execute(query,row)
                    cursor.commit()
                except:
                    pass
    def importfromtxt(self, txtpath, table_name, delimiter1,delimiter2=None):
        """
        need to preset destination table with right data type otherwise data type mismatch will be raised.
        speed of this function is at around 10'000 record per second.
        """
        cursor = self.cursor

        with open(txtpath) as f:

            values=f.readlines()[:1]

            length=len(str(values[0]).split(delimiter1))
            
        query= 'insert into '+table_name+' values ({0})'
        query=query.format(','.join('?'*length))
        
        with open(txtpath) as f:
            next(f)
            for row in f:  
                #second delimiter if seprated text with quotation marks
                
                if delimiter2 is not None:
                    
                    r=str(row).replace(delimiter2,'')
                
                    cursor.execute(query,r.split(delimiter1))

                else:

                    cursor.execute(query,row.split(delimiter1))

                cursor.commit()

     
if __name__=='__main__':
 #testing
    path=r'C:\Users\hgao1995\Desktop\actual'

    dbpath=path+'\\test.accdb'

    txt=path+'\\MKT-201903.txt'
    print(txt)
    db=Access(dbpath)

    #db.execute('Create Table test (FCNB Number, BusDate datetime, MktSegDesc char, '
                #'Sold number, USDRev number, LCRev number, USADR number, LCADR number)')

    db.importfromtxt(txt, 'test', ',','"')

    db.closecurrentdb()
