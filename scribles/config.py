#coding=utf-8

import xml.dom.minidom
import os
import sys

def get_conf(path):
	dom = xml.dom.minidom.parse(path)

	root = dom.documentElement
	url = root.getElementsByTagName('url')[0].getAttribute('value')
	host = root.getElementsByTagName('hostdomain')[0].getAttribute('value')
	path = root.getElementsByTagName('path')[0].getAttribute('value')
	filename = root.getElementsByTagName('filename')[0].getAttribute('value')
	filex = root.getElementsByTagName('filex')[0].getAttribute('value')
	conf_dict = {'url':url,'host':host,'path':path,'filename':filename,'filex':filex}
	return conf_dict

#get_conf('c:\\users\\aldin\\desktop\\scribles\\demo_config.xml')
