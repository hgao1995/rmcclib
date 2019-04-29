import paramiko


    def connect(from_dir,to_dir,download=False):
        
    
        host="www.hiltonsites.com"
        user="RMCCAPAC"
        pswd="AP@C-$f+p"

        transport = paramiko.Transport((host, 22))

        transport.connect(username = user, password = pswd)

        sftp = paramiko.SFTPClient.from_transport(transport)

        sftp.chdir(from_dir)

        if download=True:

            for file in sftp.listdir(): 

                print(file)

                localpath=to_dir+'\\'+file

                sftp.get(file,localpath)
                
        else:

            for file in sftp.listdir(): 

                print(file)            
