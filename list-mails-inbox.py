#!/usr/bin/python

# This script was inspired from: https://github.com/kautsig/ews-orgmode

import os
from lxml import etree
import httplib
import base64
import ConfigParser

# Read the config file
config = ConfigParser.RawConfigParser()
dir = os.path.realpath(__file__)[:-19]
config.read(dir + 'config.cfg')

# Exchange user and password
ewsHost = config.get('ews-mail', 'host')
ewsUrl = config.get('ews-mail', 'path')
ewsUser = config.get('ews-mail', 'username')
ewsPassword = config.get('ews-mail', 'password')
timezoneLocation = config.get('ews-mail', 'timezone')

request = """<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <FindItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
              Traversal="Shallow">
      <ItemShape>
        <t:BaseShape>Default</t:BaseShape>
      </ItemShape>
      <ParentFolderIds>
        <t:DistinguishedFolderId Id="inbox"/>
      </ParentFolderIds>
    </FindItem>
  </soap:Body>
</soap:Envelope>""".format()

# Build authentication string, remove newline for using it in a http header
auth = base64.encodestring("%s:%s" % (ewsUser, ewsPassword)).replace('\n', '')
conn = httplib.HTTPSConnection(ewsHost)
conn.request("POST", ewsUrl, body = request, headers = {
  "Host": ewsHost,
  "Content-Type": "text/xml; charset=UTF-8",
  "Content-Length": len(request),
  "Authorization" : "Basic %s" % auth
})

# Read the webservice response
resp = conn.getresponse()
data = resp.read()
conn.close()

# Debug code
print data
# exit(0)

# Parse the result xml
root = etree.fromstring(data)

xpathStr = "/s:Envelope/s:Body/m:FindItemResponse/m:ResponseMessages/m:FindItemResponseMessage/m:RootFolder/t:Items/t:Message"
namespaces = {
    's': 'http://schemas.xmlsoap.org/soap/envelope/',
    't': 'http://schemas.microsoft.com/exchange/services/2006/types',
    'm': 'http://schemas.microsoft.com/exchange/services/2006/messages',
}

# Print Mail properties
elements = root.xpath(xpathStr, namespaces=namespaces)
for element in elements:
  subject=element.find('{http://schemas.microsoft.com/exchange/services/2006/types}Subject').text
  msg_id= element.find('{http://schemas.microsoft.com/exchange/services/2006/types}ItemId').attrib['Id']
  size= element.find('{http://schemas.microsoft.com/exchange/services/2006/types}Size').text
  sensitivity= element.find('{http://schemas.microsoft.com/exchange/services/2006/types}Sensitivity').text
  print "* Subject " + subject.encode('ascii', 'ignore')
  print "* Message ID " + msg_id.encode('ascii', 'ignore')
  print "* Size " + size.encode('ascii', 'ignore')
  print "* Sensitivity " + sensitivity.encode('ascii', 'ignore')
