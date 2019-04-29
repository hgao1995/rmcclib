"Rewrite Date format according to VBA format"
#Created by Henry Gao on the 10th Jan 2019
#Modified by Henry Gao on the 11th Jan 2019

from datetime import date,datetime,timedelta
#create a dictionary that indicates format type
trans_dic={}
trans_dic['yyyymmdd']='%Y%m%d'
trans_dic['dd/mm/yyyy']='%d/%m/%Y'
trans_dic['mm/dd/yyyy']='%m/%d/%Y'
trans_dic['yyyy']='%Y'
trans_dic['mm']='%m'
trans_dic['mmm']='%b'
trans_dic['dd']='%d'
trans_dic['ddd']='%a'

def rmccformat(date_value,format_type):
    if date_value=='Today':
        result_type=trans_dic[format_type]
        return datetime.now().strftime(result_type)
    else:
        date=datetime.strptime(date_value,'%d/%m/%Y')
        result_type=trans_dic[format_type]
        return date.strftime(result_type)


#dateadd function for the moment cannot accumulate month and year
def dateadd(amount,cri,date_value):
        if date_value=='Today':
            date_a=datetime.now()
            if cri=='d':
                delta=timedelta(days=amount)
                date_b=date_a+delta
            if cri=='m':
                dateserial=date_a.strftime('%Y%m%d')
                date_b=date(year=int(dateserial[0:4]),month=int(int(dateserial[5:6])+amount),day=int(dateserial[-2:]))
            if cri=='y':
                dateserial=date_a.strftime('%Y%m%d')
                date_b=date(year=int(dateserial[0:4])+amount,month=int(dateserial[5:6]),day=int(dateserial[-2:]))                
            return date_b.strftime('%d/%m/%Y')
        else:
            date_a=datetime.strptime(date_value,'%d/%m/%Y')
            if cri=='d':
                delta=timedelta(days=amount)
                date_b=date_a+delta
            if cri=='m':
                dateserial=date_a.strftime('%Y%m%d')
                date_b=date(year=int(dateserial[0:4]),month=int(int(dateserial[5:6])+amount),day=int(dateserial[-2:]))
            if cri=='y':
                dateserial=date_a.strftime('%Y%m%d')
                date_b=date(year=int(dateserial[0:4])+amount,month=int(dateserial[5:6]),day=int(dateserial[-2:]))                
            return date_b.strftime('%d/%m/%Y')
        

