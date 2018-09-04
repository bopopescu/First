import uuid
#import hashlib
import time
import base64
import os
def get_file_name(dir, file_extension):
    f_list = os.listdir(dir)

    result_list = []

    for file_name in f_list:
        if os.path.splitext(file_name)[1] == file_extension:
            result_list.append(os.path.join(dir, file_name))
    return result_list


def get_mac_address(): 
    mac=uuid.UUID(int = uuid.getnode()).hex[-12:] 
    return mac
    #return ":".join([mac[e:e+2] for e in range(0,11,2)])



def createLicense(mac,deadLine):
    name = mac+'_license.lic'

    str1 = 'Name:'+mac
    #str1 = str1.encode(encoding="utf-8")
    strb2 = base64.encodestring(bytes(mac,'utf8'))
    str2 = strb2.decode()
    strb3 = base64.encodestring(bytes(deadLine,'utf8'))
    str3 = strb3.decode()
    #string = str(str1+str2+str3)
    fp = open(name, 'w')
    strList = [str1,str2,str3]
    for item in strList:
        fp.writelines(item)
        fp.write('\t')
        #fp.write('\n')
    fp.close

def checkLicense():
    second = time.time()
    cwd = os.getcwd()
    mac = get_mac_address()
    licenses = get_file_name(cwd, '.lic')[0]

    file_object = open(licenses) 
    try:
        file_context = file_object.read() 

    finally:
        file_object.close()
    contents=file_context.split('\t')
    [name, macCode,deadLineCode] = contents[0:3]
    deadLineBytes = base64.decodestring(deadLineCode.encode())
    macBytes = base64.decodestring(macCode.encode())
    macL = macBytes.decode()
    deadLine = float(deadLineBytes.decode())
    if (mac == macL)&(second<=deadLine) :# ok
        print('ok')
    else:
        print ('pls update license')


# generate license
mac = get_mac_address()
deadLine = str(time.mktime(time.strptime("2019-01-01","%Y-%m-%d")))  #deadline date
createLicense(mac,deadLine)

# interpret license

# m = hashlib.md5()
# m.update(mac.encode('utf8'))
# macMD5 = m.hexdigest()


checkLicense()