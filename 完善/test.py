import configparser as cp
import os

config = cp.RawConfigParser()
file_path, code_file = os.path.split(os.path.realpath(__file__))

config.read(file_path+'\\sf_config.ini', encoding='utf-8')

sf_config = config['DEFAULT']
path_project = sf_config['project']


config.read(path_project+'\\sender.ini', encoding='utf-8')
print(config.sections())
sender_config = config['DEFAULT']
sender, sig_extra = [], []
index = ['name', 'address', 'password', 'smtp_server', 'smtp_port', 'sig_pic']
for j in sender_config:
    print(j)
    if j not in index:
        print(sender_config[j])
        sig_extra.append(sender_config[j].replace('\n', '<br />'))
print(sig_extra)
