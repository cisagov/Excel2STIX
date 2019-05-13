"""
    excel2stix.py EXCEL_FILENAME

    This script will take an Excel spreadsheet and output a STIX XML
    document. The Microsoft Excel spreadsheet contains one sheet per
    indicator type with labels of:  URL, IPv4, Link, File, E-mail, FQDN,
    User Agent, Mutex, Registry, and Network Connection.  Worksheets
    can be omitted.  You can have a worksheet named Main containing
    overall header metadata. The output will have the same path/filename
    but an `.xml' file extension.


"""
import os
import sys
import openpyxl
import pprint
import time
import uuid
import codecs
import warnings
import stix.common
import stix.common.kill_chains
import stix.core
import stix.indicator
import cybox
from stix.common.vocabs import VocabString


#------------------------- common code ------------------------

# # Retrieves the namespace_Tag from the excel file input
# # @return Namespace_Tag variable as a string
# def getNamespace():
#     try:
#         filename = sys.argv[1]
#         if os.path.isfile(filename) and os.access(filename, os.R_OK):
#             wb = openpyxl.load_workbook(filename=sys.argv[1])
#         else:
#             print "Can't open the excel file '"+filename+"'"
#             sys.exit(0)
#     row = 2
#     namespace_Tag = str(sheet_ranges['H'+str(row)].value).strip()
#     print ("Namespace_Tag =" +  namespace_Tag)
#     exit()

# Handle special characters in strings
# @param string - Input string with possible unicode characters
# @return normalized unicode (best format w/i stix api)
def fix(string):
    try:
        if string is not None:
            string = str(string)
            if isinstance(string, str):
                string = string.strip()
                string = unicode(string, "utf-8")   #Make unicode
            #print "zzzzzzzz ",type(string)
    except:
        string = ""
    return string


# The below is common code used by several different scripts.
# If the below is changed, update the other scripts

#************
# Create the stix header block
# Dependencies:  import datetime, time
# @param date - Input string in form 2016-12-30T12:00:00
# @param title - Input string
# @param intent - Input string like "Indictors" or "Indicators - Watchlist"
# @param color - Input string, "WHITE","GREEN","AMBER", or "RED"
# @param fouo - Input string in form "True" or "False"
# @param desc - Input string
# @return header JSON block read from the 'Main' worksheet tab
def getHeader(date, title, intent, color, fouo, desc):
    header = {}

    if len(date) > 10:
        date = date.replace(" ","T")
    else:
       now = datetime.datetime.now()
       date = str(now.year)+"-"+str(now.month)+"-"+str(now.day)+"T"+\
            str(now.hour)+":"+str(now.minute)+":"+str(now.second)
    produced_time = {}
    produced_time["produced_time"] = date
    time = {}
    time["time"] = produced_time
    information_source = {}
    header["information_source"] = time

    header["title"] =  title

    if intent != None and intent.upper() != "NONE" and intent != "":
        intents = {}
        intents["value"] = intent
        intents["xsi:type"] = "stixVocabs:PackageIntentVocab-1.0"
        package_intents = []
        package_intents.append(intents)
        header["package_intents"] = package_intents

    marking_structures = getMarkingStructure(color)

    if fouo.upper() == "TRUE":
        tou = {"xsi:type":"TOUMarking:TermsOfUseMarkingStructureType"}
        tou["terms_of_use"] = 'WARNING: This document is FOR OFFICIAL USE ONLY (FOUO). It contains information that may be exempt from public release under the Freedom of Information Act (5 U.S.C. 552). It is to be controlled, stored, handled, transmitted, distributed, and disposed of in accordance with DHS policy relating to FOUO information and is not to be released to the  public or other personnel who do not have a valid "need-to-know" without prior approval of an authorized DHS official.'
        marking_structures.append(tou)
    handlingDict = {}
    handlingDict["marking_structures"] = marking_structures
    handlingDict["controlled_structure"] = "//node() | @*"
    handling = []
    handling.append(handlingDict)
    header["handling"] = handling
    if desc != None and desc.upper() != "NONE" and desc != "":
        header["description"]=desc
    return header
    #End getHeader


#************
# Create the JSON TLP Markings block
# @param color String WHITE, GREEN, AMBER, RED, or NONE
# @return the MarkingStructure list (may return [])
def getMarkingStructure(color):
    marking_structures = []
    if color != None and color.upper() != "NONE":
        tlpBlock = {}
        tlpBlock["color"] = color
        tlpBlock["xsi:type"] = "tlpMarking:TLPMarkingStructureType"
        marking_structures.append(tlpBlock)
        statement = {"xsi:type":"TOUMarking:TermsOfUseMarkingStructureType"}
        tou = 'DISCLAIMER: This report is provided "as is" for informational '
        tou=tou+'purposes only. The Department of Homeland Security (DHS) does'
        tou=tou+' not provide any warranties of any kind regarding any '
        tou=tou+'information contained within. The DHS does not endorse any '
        tou=tou+'commercial product or service, referenced in this bulletin or '
        tou=tou+'otherwise. '
        if color.upper() == "RED":
            tou=tou+'This document is distributed as TLP:RED: '
            tou=tou+'Not for disclosure, restricted to participants only. '
            tou=tou+'Recipients may not share TLP:RED information with any '
            tou=tou+'parties outside of their specific exchange, meeting, or '
            tou=tou+'conversation in which it was originally disclosed. In the '
            tou=tou+'context of a meeting, for example, TLP:RED information is '
            tou=tou+'limited to those present at the meeting. In most '
            tou=tou+'circumstances, TLP:RED should be exchanged verbally or in '
            tou=tou+'person. '
        elif color.upper() == "AMBER":
            tou=tou+'This document is distributed as TLP:AMBER: '
            tou=tou+"Limited disclosure, restricted to participants' "
            tou=tou+'organizations. Recipients may only use TLP:AMBER'
            tou=tou+' information with members of their own organization, and '
            tou=tou+'with clients or customers who need to know the information'
            tou=tou+' to protect themselves or prevent further harm. '
        elif color.upper() == "GREEN":
            tou=tou+'This document is distributed as TLP:GREEN: '
            tou=tou+'Limited disclosure, restricted to the community. '
            tou=tou+'Recipients may share TLP:GREEN information with peers and '
            tou=tou+'partner organizations within their sector or community, '
            tou=tou+'but not via publicly accessible channels. Information in '
            tou=tou+'this category can be circulated widely within a particular'
            tou=tou+' community. TLP:GREEN information may not be released '
            tou=tou+'outside of the community. '
        elif color.upper() == "WHITE":
            tou=tou+'This document is distributed as TLP:WHITE: '
            tou=tou+'Disclosure is not limited. '
        tou=tou+'For more information on the Traffic Light Protocol, '
        tou=tou+'see http://www.us-cert.gov/tlp.'
        statement["terms_of_use"] = tou
        marking_structures.append(statement)
    return marking_structures
    #end getMarkingStructure


#************
# Create a Network Connection Indicator Block
# @param source - input String e.g., "124.10.103.0"
# @param sspoofed - input String, either "True" or "False"
# @param sport - input String the source port e.g., "8080"
# @param sproto - input String either "UDP" or "TCP"
# @param dest - input String, e.g., "250.0.0.1"
# @param dspoofed - input String either "True" or "False"
# @param dport - input String the destination port e.g., "53127"
# @param dproto - input String either "UDP" or "TCP"
# @returns prop - output JSON dictionary
def getNetConn(source,sspoofed, sport, sproto, dest, dspoofed, dport, dproto):

    prop = {}

    item = None
    srcaddr = {"xsi:type":"SocketAddressObjectType"}
    if source != None and source != '' and source.upper() != "NONE":
        item = {"condition":"Equals"}
        item["value"]= source
        ipaddr={"category":"ipv4-addr","xsi:type":"AddressObjectType","is_source":True}
        if sspoofed.upper() == "TRUE":
            ipaddr["is_spoofed"] = True
        else:
            ipaddr["is_spoofed"] = False
        srcaddr["ip_address"] = ipaddr
        if item != None:
            ipaddr["address_value"] = item

    if sport != None and sport != '' and sport.upper() != "NONE":
        layer4_proto = None
        if sproto != None and sproto != '' and sproto.upper() != "NONE":
            layer4_proto={"condition":"Equals","is_obfuscated":True}
            layer4_proto["value"] = sproto
        port_block = {"xsi:type":"PortObjectType"}
        if layer4_proto != None:
            port_block["layer4_protocol"] = layer4_proto
        item = {"condition":"Equals"}
        item["value"] = int(sport)
        port_block["port_value"] = item
        srcaddr["port"] = port_block
        prop["source_socket_address"] = srcaddr

    item = None
    dstaddr = {"xsi:type":"SocketAddressObjectType"}
    if dest != None and dest != '' and dest.upper() != "NONE":
        item = {"condition":"Equals"}
        item["value"]= dest
        ipaddr={"category":"ipv4-addr","xsi:type":"AddressObjectType","is_source":False}
        if dspoofed.upper() == "TRUE":
            ipaddr["is_spoofed"] = True
        else:
            ipaddr["is_spoofed"] = False
        dstaddr["ip_address"] = ipaddr
        if item != None:
            ipaddr["address_value"] = item

    if dport != None and dport != '' and dport.upper() != "NONE":
        layer4_proto = None
        if dproto != None and dproto != '' and dproto.upper() != "NONE":
            layer4_proto={"condition":"Equals","is_obfuscated":True}
            layer4_proto["value"] = dproto
        port_block = {"xsi:type":"PortObjectType"}
        if layer4_proto != None:
            port_block["layer4_protocol"] = layer4_proto
        item = {"condition":"Equals"}
        item["value"] = int(dport)
        port_block["port_value"] = item
        dstaddr["port"] = port_block
        prop["destination_socket_address"] = dstaddr

    if prop != {}:
        prop["xsi:type"] = "NetworkConnectionObjectType"

    return prop
    #end getNetConn


#************
# Create the JSON sighting block
# Dependencies:  import datetime, time
# @param datetime of the form "2015-10-28T12:00:00" or "2015-10-28 12:00:00"
# @return sighting dictionary
def getSightings(dateStr):
    dateStr = dateStr.strip()
    if len(dateStr) < 10:
        now = datetime.datetime.now()
        dateStr = str(now.year)+"-"+str(now.month)+"-"+\
            str(now.day)+"T"+str(now.hour)+":"+\
              str(now.minute)+":"+str(now.second)
    elif len(dateStr) == 10:
        dateStr = dateStr + "T00:00:00"
    else:
        dateStr = dateStr.replace(" ","T")
    dateStr = dateStr.upper()
    sighting = {"timestamp_precision":"second"}
    sighting["timestamp"] = dateStr
    sightList = []
    sightList.append(sighting)
    sightDict = {}
    sightDict["sightings"] = sightList
    sightDict["sightings_count"] = 1
    return sightDict
    #end getSightings

#------------------------ end common code use ----------------------------

class excel2stix():

    # Constructor for class
    def __init__(self):
        self.DEBUG = False  #MKPMKP
        self.__version__ = "0.1y"
        self.indicators = []


    # Get the stix version
    # @returns<string> stix version, e.g., "1.1.1.0"
    def getStixVersion(self):
        return stix.__version__

    # Get the cybox version
    # @returns<string> cybox version, e.g., "2.1.0.4"
    def getCyboxVersion(self):
        return cybox.__version__

    # Get the excel2stix version
    # @returns<string> excel2stix version, e.g., "1.0c"
    def getVersion(self):
        return self.__version__

    # Check string to see if a string is null
    # @param token - Input string to check
    # @return True if null (or empty)
    def isNull(self, token):
        flag = False
        if token is None:
            flag = True
        else:
            token = (token.strip()).upper()
            if token == 'NONE' or token == '':
                flag = True
        return flag


    # Set the JSON structure with the default killchain definitions
    # @return  Dictionary containing killchain definitions
    def setKillChains(self):
	killphase = []
	killphase.append({"name":"Reconnaissance","ordinality":1,"phase_id":
            "stix:KillChainPhase-af1016d6-a744-4ed7-ac91-00fe2272185a"})
	killphase.append({"name":"Weaponization","ordinality":2,"phase_id":
            "stix:KillChainPhase-445b4827-3cca-42bd-8421-f2e947133c16"})
	killphase.append({"name":"Delivery","ordinality":3,"phase_id":
            "stix:KillChainPhase-79a0e041-9d5f-49bb-ada4-8322622b162d"})
	killphase.append({"name":"Exploitation","ordinality":4,"phase_id":
            "stix:KillChainPhase-f706e4e7-53d8-44ef-967f-81535c9db7d0"})
	killphase.append({"name":"Installation","ordinality":5,"phase_id":
            "stix:KillChainPhase-e1e4e3f7-be3b-4b39-b80a-a593cfd99a4f"})
	killphase.append({"name":"Command and Control","ordinality":6,
	    "phase_id":
	    "stix:KillChainPhase-d6dc32b9-2538-4951-8733-3cb9ef1daae2"})
	killphase.append({"name":"Actions on Objectives","ordinality":7,
	    "phase_id":
	    "stix:KillChainPhase-786ca8f9-2d9a-4213-b38e-399af4a2e5d6"})
	killchains = {"definer":"LMCO","id":
	    "stix:KillChainPhase-af3e707f-2fb9-49e5-8c37-14026ca0a5ff",
	    "name": "LM Cyber Kill Chain",
	    "number_of_phases": "7",
	    "reference":
            "http://www.lockheedmartin.com/content/dam/lockheed/data/corporate/documents/LM-White-Paper-Intel-Driven-Defense.pdf"}
	killchains["kill_chain_phases"] = killphase
	killList = []
	killList.append(killchains)
	killBlock = {}
	killBlock["kill_chains"] = killList
	killSection = {}
	killSection["kill_chains"] = killBlock
        return killSection



    # Create the common indicator JSON block
    # @param title String like "Malicious FQDN Indicator"
    # @param desc String, user-defined
    # @param color String WHITE or GREEN or AMBER or RED
    # @param type String such as Anonymization or C2
    # @param fouo String TRUE or FALSE
    # @param sighted String with date/time
    # @param killphase String such as Installation or Exploitation
    # @return dictionary JSON Indicator block
    def doCommon(self, title, desc, color, type, fouo, sighted, killphase):
        if self.DEBUG:
            print '--- ',title,' ---'
            print 'desc=     ',desc
            print 'type=     ',type
            print 'tlp=      ',color
            print 'fouo=     ',fouo
            print 'sighted=  ',sighted
            print 'killphase=',killphase
        ind = {}
        ind["id"] = namespace_Tag+":indicator-"+str(uuid.uuid1())
        ind["title"] = title
        if self.isNull(desc) != True:
            ind["description"] = desc

	marking_structures = getMarkingStructure(color)

        if fouo.upper() == "TRUE":
            tou = {"xsi:type":"TOUMarking:TermsOfUseMarkingStructureType"}
            tou["terms_of_use"] = 'WARNING: This document is FOR OFFICIAL USE ONLY (FOUO). It contains information that may be exempt from public release under the Freedom of Information Act (5 U.S.C. 552). It is to be controlled, stored, handled, transmitted, distributed, and disposed of in accordance with DHS policy relating to FOUO information and is not to be released to the  public or other personnel who do not have a valid "need-to-know" without prior approval of an authorized DHS official.'
            marking_structures.append(tou)

        handlingDict = {}
        if marking_structures != []:
            handlingDict["marking_structures"] = marking_structures
        handling = []

        # Suppress empty handling blocks
        if handlingDict:
            handling.append(handlingDict)
            ind["handling"] = handling

        if self.isNull(type) != True:
            if type == "Benign" or type == "Compromised":
                indicatorList = [type]
            else:
                indicatorDict={"xsi:type":"stixVocabs:IndicatorTypeVocab-1.1"}
                indicatorDict["value"]=type
                indicatorList = []
                indicatorList.append(indicatorDict)
            ind["indicator_types"] = indicatorList
        if self.isNull(sighted) != True:
            ind["timestamp"] = sighted
            ind["sightings"] = getSightings(sighted)

        killstr = {"kill_chain_id":"stix:KillChainPhase-af3e707f-2fb9-49e5-8c37-14026ca0a5ff"}
        if killphase == "RECONNAISSANCE":
            killstr["phase_id"] = \
                "stix:KillChainPhase-af1016d6-a744-4ed7-ac91-00fe2272185a"
            killstr["ordinality"] = "1"
            killstr["name"] = "Reconnaissance"
        elif killphase == "WEAPONIZATION":
            killstr["phase_id"] = \
                "stix:KillChainPhase-445b4827-3cca-42bd-8421-f2e947133c16"
            killstr["ordinality"] = "2"
            killstr["name"] = "Weaponization"
        elif killphase == "DELIVERY":
            killstr["phase_id"] = \
                "stix:KillChainPhase-79a0e041-9d5f-49bb-ada4-8322622b162d"
            killstr["ordinality"] = "3"
            killstr["name"] = "Delivery"
        elif killphase == "EXPLOITATION":
            killstr["phase_id"] =\
                "stix:KillChainPhase-f706e4e7-53d8-44ef-967f-81535c9db7d0"
            killstr["ordinality"] = "4"
            killstr["name"] = "Exploitation"
        elif killphase == "INSTALLATION":
            killstr["phase_id"] =\
                "stix:KillChainPhase-e1e4e3f7-be3b-4b39-b80a-a593cfd99a4f"
            killstr["ordinality"] = "5"
            killstr["name"] = "Installation"
        elif killphase == "COMMAND AND CONTROL":
            killstr["phase_id"] =\
                "stix:KillChainPhase-d6dc32b9-2538-4951-8733-3cb9ef1daae2"
            killstr["ordinality"] = "6"
            killstr["name"] = "Command and Control"
        elif killphase == "ACTIONS ON OBJECTIVES":
            killstr["phase_id"] =\
                "stix:KillChainPhase-786ca8f9-2d9a-4213-b38e-399af4a2e5d6"
            killstr["ordinality"] = "7"
            killstr["name"] = "Actions on Objectives"
        if 'phase_id' in killstr:
            killList = []
            killList.append(killstr)
            killDict = {}
            killDict["kill_chain_phases"] = killList
            ind["kill_chain_phases"] = killDict
        return ind


    # Create an URI/URL JSON Indicator block
    # @param desc String description
    # @param type String: Malicious E-Mail, IP Watchlist, Domain Watchlist,
    #           URL Watchlist, Malware Artifacts, C2, Anonymization,
    #           Exfiltration, Host Characteristics, File Hash Watchlist
    # @param color String: WHITE, GREEN, AMBER, RED
    # @param fouo String: "TRUE" or "FALSE"
    # @param sighted String: of the form "2015-12-30T12:00:00"
    # @parapm killphase String: Reconnaissance, Weaponization, Delivery,
    #           Exploitation, Installation, Command and Control, Actions on
    #           Objectives
    def doUrl(self, desc, type, color, fouo, sighted, killphase, url):

        ind = self.doCommon("Malicious URL Indicator",desc,color,type,\
            fouo, sighted,killphase)
        prop = {"type":"URL","xsi:type":"URIObjectType"}
        urlValue = {"condition":"Equals"}
        urlValue["value"] = url
        prop["value"] = urlValue
        obj = {}
        obj["id"] = namespace_Tag+":Object-"+str(uuid.uuid1())
        obj["properties"] = prop
        observable = {}
        observable["id"] = namespace_Tag+":Observable-"+str(uuid.uuid1())
        observable["object"] = obj
        ind["observable"] = observable
        return ind



    # Create an FQDN JSON Indicator block
    # @param desc String description
    # @param type String: Malicious E-Mail, IP Watchlist, Domain Watchlist,
    #           URL Watchlist, Malware Artifacts, C2, Anonymization,
    #           Exfiltration, Host Characteristics, File Hash Watchlist
    # @param color String: WHITE, GREEN, AMBER, RED
    # @param fouo String: "TRUE" or "FALSE"
    # @param sighted String: of the form "2015-12-30T12:00:00"
    # @parapm killphase String: Reconnaissance, Weaponization, Delivery,
    #           Exploitation, Installation, Command and Control, Actions on
    #           Objectives
    def doFqdn(self, desc, type, color, fouo, sighted, killphase, fqdn):

        ind = self.doCommon("Malicious FQDN Indicator",desc,color,type,\
            fouo, sighted,killphase)
        prop = {"type":"FQDN","xsi:type":"DomainNameObjectType"}
        fqdnValue = {"condition":"Equals"}
        fqdnValue["value"] = fqdn
        prop["value"] = fqdnValue
        obj = {}
        obj["id"] = namespace_Tag+":Object-"+str(uuid.uuid1())
        obj["properties"] = prop
        observable = {}
        observable["id"] = namespace_Tag+":Observable-"+str(uuid.uuid1())
        observable["object"] = obj
        ind["observable"] = observable
        return ind


    # Create a Mutex JSON Indicator block
    # @param desc String description
    # @param type String: Malicious E-Mail, IP Watchlist, Domain Watchlist,
    #           URL Watchlist, Malware Artifacts, C2, Anonymization,
    #           Exfiltration, Host Characteristics, File Hash Watchlist
    # @param color String: WHITE, GREEN, AMBER, RED
    # @param fouo String: "TRUE" or "FALSE"
    # @param sighted String: of the form "2015-12-30T12:00:00"
    # @parapm killphase String: Reconnaissance, Weaponization, Delivery,
    #           Exploitation, Installation, Command and Control, Actions on
    #           Objectives
    def doMutex(self, desc, type, color, fouo, sighted, killphase, mutex):

        ind = self.doCommon("Malicious Mutex Indicator",desc,color,type,\
            fouo, sighted,killphase)
        prop = {"type":"Mutex","xsi:type":"MutexObjectType"}
        name = {"condition":"Equals"}
        name["value"] = mutex
        #prop["condition"] = "Equals"
        prop["name"] = name
        obj = {}
        obj["id"] = namespace_Tag+":Object-"+str(uuid.uuid1())
        obj["properties"] = prop
        observable = {}
        observable["id"] = namespace_Tag+":Observable-"+str(uuid.uuid1())
        observable["object"] = obj
        ind["observable"] = observable
        return ind

    # Create a User Agent Indicator Block
    # @param desc String description
    # @param type String: Malicious E-Mail, IP Watchlist, Domain Watchlist,
    #           URL Watchlist, Malware Artifacts, C2, Anonymization,
    #           Exfiltration, Host Characteristics, File Hash Watchlist
    # @param color String: WHITE, GREEN, AMBER, RED
    # @param fouo String: "TRUE" or "FALSE"
    # @param sighted String: of the form "2015-12-30T12:00:00"
    # @parapm killphase String: Reconnaissance, Weaponization, Delivery,
    #           Exploitation, Installation, Command and Control, Actions on
    #           Objectives
    def doUa(self, desc, type, color, fouo, sighted, killphase, ua):

        ind = self.doCommon("Malicious User Agent Indicator",desc,color,type,\
            fouo, sighted,killphase)
        prop = {"type":"User Agent","xsi:type":"HTTPSessionObjectType"}
        inner = {"condition":"Equals"}
        inner["value"] = ua
        header = {}
        header["user_agent"] = inner
        middle = {}
        middle["parsed_header"] = header
        outer = {}
        outer["http_request_header"] = middle
        request = {}
        request["http_client_request"] = outer
        response = []
        response.append(request)
        prop["http_request_response"] = response

        obj = {}
        obj["id"] = namespace_Tag+":Object-"+str(uuid.uuid1())
        obj["properties"] = prop
        observable = {}
        observable["id"] = namespace_Tag+":Observable-"+str(uuid.uuid1())
        observable["object"] = obj
        ind["observable"] = observable
        return ind


    # Create a Windows Registry Indicator Block
    # @param desc String description
    # @param type String: Malicious E-Mail, IP Watchlist, Domain Watchlist,
    #           URL Watchlist, Malware Artifacts, C2, Anonymization,
    #           Exfiltration, Host Characteristics, File Hash Watchlist
    # @param color String: WHITE, GREEN, AMBER, RED
    # @param fouo String: "TRUE" or "FALSE"
    # @param sighted String: of the form "2015-12-30T12:00:00"
    # @parapm killphase String: Reconnaissance, Weaponization, Delivery,
    #           Exploitation, Installation, Command and Control, Actions on
    #           Objectives
    def doRegistry(self, desc, type, color, fouo, sighted, killphase,\
       hive, key, name, data ):

        ind = self.doCommon("Malicious Registry Indicator",desc,color,type,\
            fouo, sighted,killphase)
        prop = {"type":"WIndows Registry","xsi:type":"WindowsRegistryKeyObjectType"}
        item = {"condition":"Equals"}
        item["value"] = hive
        prop["hive"] = item
        if self.isNull(key) != True:
            item = {"condition":"Equals"}
            item["value"] = key
            prop["key"] = item
        inner = []
        values = {}
        if self.isNull(data) != True:
            item = {"condition":"Equals"}
            item["value"] = data
            values["data"] = item
        if self.isNull(name) != True:
            item = {"condition":"Equals"}
            item["value"] = name
            values["name"] = item
        inner.append(values)
        prop["values"] = inner

        obj = {}
        obj["id"] = namespace_Tag+":Object-"+str(uuid.uuid1())
        obj["properties"] = prop
        observable = {}
        observable["id"] = namespace_Tag+":Observable-"+str(uuid.uuid1())
        observable["object"] = obj
        ind["observable"] = observable
        return ind

    # Create a link Indicator Block
    # @param desc String description
    # @param type String: Malicious E-Mail, IP Watchlist, Domain Watchlist,
    #           URL Watchlist, Malware Artifacts, C2, Anonymization,
    #           Exfiltration, Host Characteristics, File Hash Watchlist
    # @param color String: WHITE, GREEN, AMBER, RED
    # @param fouo String: "TRUE" or "FALSE"
    # @param sighted String: of the form "2015-12-30T12:00:00"
    # @parapm killphase String: Reconnaissance, Weaponization, Delivery,
    #           Exploitation, Installation, Command and Control, Actions on
    #           Objectives
    def doLink(self, desc, type, color, fouo, sighted, killphase, link, name):

        ind = self.doCommon("Malicious Link Indicator",desc,color,type,\
            fouo, sighted,killphase)
        prop = {"type":"URL","xsi:type":"LinkObjectType"}
        linkValue = {"condition":"Equals"}
        linkValue["value"] = link
        prop["value"] = linkValue
        prop["condition"] = "Equals"
        label = {"condition":"Equals"}
        label["value"] = name
        prop["url_label"] = label

        obj = {}
        obj["id"] = namespace_Tag+":Object-"+str(uuid.uuid1())
        obj["properties"] = prop
        observable = {}
        observable["id"] = namespace_Tag+":Observable-"+str(uuid.uuid1())
        observable["object"] = obj
        ind["observable"] = observable
        return ind

    # Create an IPv4 Indicator Block
    # @param desc String description
    # @param type String: Malicious E-Mail, IP Watchlist, Domain Watchlist,
    #           URL Watchlist, Malware Artifacts, C2, Anonymization,
    #           Exfiltration, Host Characteristics, File Hash Watchlist
    # @param color String: WHITE, GREEN, AMBER, RED
    # @param fouo String: "TRUE" or "FALSE"
    # @param sighted String: of the form "2015-12-30T12:00:00"
    # @parapm killphase String: Reconnaissance, Weaponization, Delivery,
    #           Exploitation, Installation, Command and Control, Actions on
    #           Objectives
    def doIpv4(self, desc, type, color, fouo, sighted, killphase, ipv4, spoofed):

        ind = self.doCommon("Malicious IPv4 Indicator",desc,color,type,\
            fouo, sighted,killphase)
        prop = {"type":"address_value","xsi:type":"AddressObjectType"}
        prop["category"] = "ipv4-addr"
        if spoofed.upper() == "TRUE":
            prop["is_spoofed"] = True
        else:
            prop["is_spoofed"] = False

        label = {}
        label["condition"] = "Equals"
        label["value"] = ipv4
        prop["address_value"] = label

        obj = {}
        obj["id"] = namespace_Tag+":Object-"+str(uuid.uuid1())
        obj["properties"] = prop
        observable = {}
        observable["id"] = namespace_Tag+":Observable-"+str(uuid.uuid1())
        observable["object"] = obj
        ind["observable"] = observable
        return ind


    # Create a File Indicator Block
    # @param desc String description
    # @param type String: Malicious E-Mail, IP Watchlist, Domain Watchlist,
    #           URL Watchlist, Malware Artifacts, C2, Anonymization,
    #           Exfiltration, Host Characteristics, File Hash Watchlist
    # @param color String: WHITE, GREEN, AMBER, RED
    # @param fouo String: "TRUE" or "FALSE"
    # @param sighted String: of the form "2015-12-30T12:00:00"
    # @parapm killphase String: Reconnaissance, Weaponization, Delivery,
    #           Exploitation, Installation, Command and Control, Actions on
    #           Objectives
    def doFile(self, desc, type, color, fouo, sighted, killphase, file, path,\
        size, md5, sha1, sha256, ssdeep):

        ind = self.doCommon("Malicious File Indicator",desc,color,type,\
            fouo, sighted,killphase)
        prop = {"type":"File","xsi:type":"FileObjectType"}
        if not (self.isNull(file)):
            nameBlock = {"condition":"Equals"}
            nameBlock["value"] = file
            prop["file_name"] = nameBlock
        if not (self.isNull(path)):
            pathBlock = {"condition":"Equals"}
            pathBlock["value"] = path
            prop["file_path"] = pathBlock
        if self.isNull(size) != True:
            sizeBlock = {"condition":"Equals"}
            sizeBlock["value"] = size
            prop["size_in_bytes"] = sizeBlock
        if not ((self.isNull(md5)) and (self.isNull(sha1)) and (self.isNull(sha256)) and
        (self.isNull(ssdeep))):
            hashes = []
            if self.isNull(md5) != True:
                hash = {}
                hash["type"] = {"condition":"Equals","value":"MD5"}
                chunk = {"condition":"Equals"}
                chunk["value"] = md5
                hash["simple_hash_value"] = chunk
                hashes.append(hash)
            if self.isNull(sha1) != True:
                hash = {}
                hash["type"] = {"condition":"Equals","value":"SHA1"}
                chunk = {"condition":"Equals"}
                chunk["value"] = sha1
                hash["simple_hash_value"] = chunk
                hashes.append(hash)
            if self.isNull(sha256) != True:
                hash = {}
                hash["type"] = {"condition":"Equals","value":"SHA256"}
                chunk = {"condition":"Equals"}
                chunk["value"] = sha256
                hash["simple_hash_value"] = chunk
                hashes.append(hash)
            if self.isNull(ssdeep) != True:
                hash = {}
                hash["type"] = {"condition":"Equals","value":"SSDEEP"}
                chunk = {"condition":"Equals"}
                chunk["value"] = ssdeep
                hash["fuzzy_hash_value"] = chunk
                hashes.append(hash)
            prop["hashes"] = hashes
        obj = {}
        obj["id"] = namespace_Tag+":Object-"+str(uuid.uuid1())
        obj["properties"] = prop
        observable = {}
        observable["id"] = namespace_Tag+":Observable-"+str(uuid.uuid1())
        observable["object"] = obj
        ind["observable"] = observable
        return ind


    # Create an E-mail Indicator Block
    # @param desc String description
    # @param type String: Malicious E-Mail, IP Watchlist, Domain Watchlist,
    #           URL Watchlist, Malware Artifacts, C2, Anonymization,
    #           Exfiltration, Host Characteristics, File Hash Watchlist
    # @param color String: WHITE, GREEN, AMBER, RED
    # @param fouo String: "TRUE" or "FALSE"
    # @param sighted String: of the form "2015-12-30T12:00:00"
    # @parapm killphase String: Reconnaissance, Weaponization, Delivery,
    #           Exploitation, Installation, Command and Control, Actions on
    #           Objectives
    def doEmail(self, desc, type, color, fouo, sighted, killphase, fromaddr,\
        spoofed, subject, msgid, xmailer):

        ind = self.doCommon("Malicious E-mail Indicator",desc,color,type,\
            fouo,sighted,killphase)
        header = {}
        if self.isNull(fromaddr) != True:
            fromBlock = {"category":"e-mail","xsi:type":"AddressObjectType"}
            item = {"condition":"Equals"}
            item["value"] = fromaddr
            fromBlock["address_value"] = item
            if spoofed.upper() == "TRUE":
                fromBlock["is_spoofed"] = True
            header["from"] = fromBlock
        if self.isNull(msgid) != True:
            item = {"condition":"Equals"}
            item["value"] = msgid
            header["message_id"] = item
        if self.isNull(xmailer) != True:
            item = {"condition":"Equals"}
            item["value"] = xmailer
            header["x_mailer"] = item
        if self.isNull(subject) != True:
            item = {"condition":"Equals"}
            item["value"] = subject
            header["subject"] = item
        prop = {}
        prop = {"xsi:type":"EmailMessageObjectType"}
        prop["header"] = header
        obj = {}
        obj["id"] = namespace_Tag+":Object-"+str(uuid.uuid1())
        obj["properties"] = prop
        observable = {}
        observable["id"] = namespace_Tag+":Observable-"+str(uuid.uuid1())
        observable["object"] = obj
        ind["observable"] = observable
        return ind



    # Create a Network Connection Indicator Block
    # @param desc String description
    # @param type String: Malicious E-Mail, IP Watchlist, Domain Watchlist,
    #           URL Watchlist, Malware Artifacts, C2, Anonymization,
    #           Exfiltration, Host Characteristics, File Hash Watchlist
    # @param color String: WHITE, GREEN, AMBER, RED
    # @param fouo String: "TRUE" or "FALSE"
    # @param sighted String: of the form "2015-12-30T12:00:00"
    # @parapm killphase String: Reconnaissance, Weaponization, Delivery,
    #           Exploitation, Installation, Command and Control, Actions on
    #           Objectives
    def doNetConn(self, desc, type, color, fouo, sighted, killphase, source,\
        sspoofed, sport, sproto, dest, dspoofed, dport, dproto):

        ind = self.doCommon("Malicious Network Connection Indicator",desc,color,type,\
            fouo, sighted,killphase)
        prop=getNetConn(source,sspoofed,sport,sproto,dest,dspoofed,dport,dproto)

        obj = {}
        obj["id"] = namespace_Tag+":Object-"+str(uuid.uuid1())
        obj["properties"] = prop
        observable = {}
        observable["id"] = namespace_Tag+":Observable-"+str(uuid.uuid1())
        observable["object"] = obj
        ind["observable"] = observable
        return ind


    # Process the excel spreadsheet
    def getIndicators(self, wb):
	indicators = []

        try:
            sheet_ranges = wb['URL']
	    row = 2
	    while True:
                desc = fix(sheet_ranges['A'+str(row)].value)
                type = str(sheet_ranges['B'+str(row)].value).strip()
                color =  str(sheet_ranges['C'+str(row)].value).strip()
	        fouo = str(sheet_ranges['D'+str(row)].value).strip()
                sighted = str(sheet_ranges['E'+str(row)].value).strip()
	        killphase=(str(sheet_ranges['F'+str(row)].value).strip()).upper()
	        url = fix(sheet_ranges['G'+str(row)].value)
	        if (self.isNull(desc)) and (self.isNull(type)) and \
                    (self.isNull(color)) and (self.isNull(fouo)) and \
                    (self.isNull(sighted)) and (self.isNull(killphase)) and \
                    (self.isNull(url)):
	                break
                row = row + 1
                indicators.append(self.doUrl(desc, type, color, fouo,\
                    sighted, killphase, url))
            print "Processed ",row-2," URL indicators"
        except KeyError:
	    print "URL Sheet does not exist!"

        try:
            sheet_ranges = wb['FQDN']
	    row = 2
	    while True:
                desc = fix(sheet_ranges['A'+str(row)].value)
                type = str(sheet_ranges['B'+str(row)].value).strip()
                color =  str(sheet_ranges['C'+str(row)].value).strip()
	        fouo = str(sheet_ranges['D'+str(row)].value).strip()
                sighted = str(sheet_ranges['E'+str(row)].value).strip()
	        killphase=(str(sheet_ranges['F'+str(row)].value).strip()).upper()
	        fqdn = fix(sheet_ranges['G'+str(row)].value)
	        if (self.isNull(desc)) and (self.isNull(type)) and \
                    (self.isNull(color)) and (self.isNull(sighted)) and \
                    (self.isNull(killphase)) and (self.isNull(fqdn)):
	                break
                row = row + 1
                indicators.append(self.doFqdn(desc, type, color, fouo,\
                    sighted, killphase, fqdn))
            print "Processed ",row-2," FQDN indicators"
        except KeyError:
	    print "FQDN Sheet does not exist!"

        try:
            sheet_ranges = wb['IPv4']
	    row = 2
	    while True:
                desc = fix(sheet_ranges['A'+str(row)].value)
                type = str(sheet_ranges['B'+str(row)].value).strip()
                color =  str(sheet_ranges['C'+str(row)].value).strip()
	        fouo = str(sheet_ranges['D'+str(row)].value).strip()
                sighted = str(sheet_ranges['E'+str(row)].value).strip()
	        killphase=(str(sheet_ranges['F'+str(row)].value).strip()).upper()
	        ipv4 =  str(sheet_ranges['G'+str(row)].value).strip()
	        spoofed =  str(sheet_ranges['H'+str(row)].value).strip()
	        if (self.isNull(desc)) and (self.isNull(type)) and \
                    (self.isNull(color)) and (self.isNull(fouo)) and \
                    (self.isNull(sighted)) and (self.isNull(killphase)) and \
                    (self.isNull(ipv4)) and (self.isNull(spoofed)):
	                break
                row = row + 1
                indicators.append(self.doIpv4(desc, type, color, fouo,\
                    sighted, killphase, ipv4, spoofed))
            print "Processed ",row-2," IPv4 indicators"
        except KeyError:
	    print "IPv4 Sheet does not exist!"

        try:
            sheet_ranges = wb['Link']
	    row = 2
	    while True:
                desc = fix(sheet_ranges['A'+str(row)].value)
                type = str(sheet_ranges['B'+str(row)].value).strip()
                color =  str(sheet_ranges['C'+str(row)].value).strip()
	        fouo = str(sheet_ranges['D'+str(row)].value).strip()
                sighted = str(sheet_ranges['E'+str(row)].value).strip()
	        killphase=(str(sheet_ranges['F'+str(row)].value).strip()).upper()
	        link = fix(sheet_ranges['G'+str(row)].value)
	        name = fix(sheet_ranges['H'+str(row)].value)
	        if (self.isNull(desc)) and (self.isNull(type)) and \
                    (self.isNull(color)) and (self.isNull(fouo)) and \
                    (self.isNull(sighted)) and (self.isNull(killphase)) and \
                    (self.isNull(link)) and (self.isNull(name)):
	                break
                row = row + 1
                indicators.append(self.doLink(desc, type, color, fouo,\
                    sighted, killphase, link, name))
            print "Processed ",row-2," Link indicators"
        except KeyError:
	    print "Link Sheet does not exist!"

        try:
            sheet_ranges = wb['File']
	    row = 2
	    while True:
                desc = fix(sheet_ranges['A'+str(row)].value)
                type = str(sheet_ranges['B'+str(row)].value).strip()
                color =  str(sheet_ranges['C'+str(row)].value).strip()
	        fouo = str(sheet_ranges['D'+str(row)].value).strip()
                sighted = str(sheet_ranges['E'+str(row)].value).strip()
	        killphase=(str(sheet_ranges['F'+str(row)].value).strip()).upper()
	        filename = fix(sheet_ranges['G'+str(row)].value)
	        filepath = fix(sheet_ranges['H'+str(row)].value)
	        filesize =  str(sheet_ranges['I'+str(row)].value).strip()
	        md5 =  str(sheet_ranges['J'+str(row)].value).strip()
	        sha1 =  str(sheet_ranges['K'+str(row)].value).strip()
	        sha256 =  str(sheet_ranges['L'+str(row)].value).strip()
	        ssdeep =  str(sheet_ranges['M'+str(row)].value).strip()
	        if (self.isNull(desc)) and (self.isNull(type)) and \
                    (self.isNull(color)) and (self.isNull(sighted)) and \
                    (self.isNull(killphase)) and (self.isNull(filename)) \
                    and (self.isNull(filepath)) and (self.isNull(filesize)) \
                    and (self.isNull(md5)) and (self.isNull(sha1)) and \
                    (self.isNull(sha256)) and (self.isNull(ssdeep)):
	                break
                row = row + 1
                indicators.append(self.doFile(desc, type, color, fouo,\
                    sighted, killphase, filename, filepath, filesize, md5,\
                    sha1, sha256, ssdeep))
            print "Processed ",row-2," file indicators"
        except KeyError:
	    print "File Sheet does not exist!"

        try:
            sheet_ranges = wb['E-mail']
	    row = 2
	    while True:
                desc = fix(sheet_ranges['A'+str(row)].value)
                type = str(sheet_ranges['B'+str(row)].value).strip()
                color =  str(sheet_ranges['C'+str(row)].value).strip()
	        fouo = str(sheet_ranges['D'+str(row)].value).strip()
                sighted = str(sheet_ranges['E'+str(row)].value).strip()
	        killphase=(str(sheet_ranges['F'+str(row)].value).strip()).upper()
	        fromaddr = fix(sheet_ranges['G'+str(row)].value)
	        spoofed =  str(sheet_ranges['H'+str(row)].value).strip()
	        subj = fix(sheet_ranges['I'+str(row)].value)
	        msgid =  fix(sheet_ranges['J'+str(row)].value)
	        xmailer =  str(sheet_ranges['K'+str(row)].value).strip()
	        if (self.isNull(desc)) and (self.isNull(type)) and \
                    (self.isNull(color)) and (self.isNull(fouo)) and \
                    (self.isNull(sighted)) and (self.isNull(killphase)) and \
                    (self.isNull(fromaddr)) and (self.isNull(spoofed)) and \
                    (self.isNull(subj)) and (self.isNull(msgid)) and \
                    (self.isNull(xmailer)):
	                break
                row = row + 1
                indicators.append(self.doEmail(desc, type, color, fouo,\
                    sighted, killphase, fromaddr, spoofed, subj, msgid,\
                    xmailer))
            print "Processed ",row-2," E-mail indicators"
        except KeyError:
	    print "E-mail Sheet does not exist!"


        try:
            sheet_ranges = wb['User Agent']
	    row = 2
	    while True:
                desc = fix(sheet_ranges['A'+str(row)].value)
                type = str(sheet_ranges['B'+str(row)].value).strip()
                color =  str(sheet_ranges['C'+str(row)].value).strip()
	        fouo = str(sheet_ranges['D'+str(row)].value).strip()
                sighted = str(sheet_ranges['E'+str(row)].value).strip()
	        killphase=(str(sheet_ranges['F'+str(row)].value).strip()).upper()
	        ua = fix(sheet_ranges['G'+str(row)].value)
	        if (self.isNull(desc)) and (self.isNull(type)) and \
                    (self.isNull(color)) and (self.isNull(fouo)) and \
                    (self.isNull(sighted)) and (self.isNull(killphase)) and \
                    (self.isNull(ua)):
	                break
                row = row + 1
                indicators.append(self.doUa(desc, type, color, fouo,\
                    sighted, killphase, ua))
            print "Processed ",row-2," User Agent indicators"
        except KeyError:
	    print "User Agent Sheet does not exist!"

        try:
            sheet_ranges = wb['Mutex']
	    row = 2
	    while True:
                desc = fix(sheet_ranges['A'+str(row)].value)
                type = str(sheet_ranges['B'+str(row)].value).strip()
                color =  str(sheet_ranges['C'+str(row)].value).strip()
	        fouo = str(sheet_ranges['D'+str(row)].value).strip()
                sighted = str(sheet_ranges['E'+str(row)].value).strip()
	        killphase=(str(sheet_ranges['F'+str(row)].value).strip()).upper()
	        mutex = fix(sheet_ranges['G'+str(row)].value)
	        if (self.isNull(desc)) and (self.isNull(type)) and \
                    (self.isNull(color)) and (self.isNull(fouo)) and \
                    (self.isNull(sighted)) and (self.isNull(killphase)) and \
                    (self.isNull(mutex)):
	                break
                row = row + 1
                indicators.append(self.doMutex(desc, type, color, fouo,\
                    sighted, killphase, mutex))
            print "Processed ",row-2," Mutex indicators"
        except KeyError:
	    print "Mutex Sheet does not exist!"

        try:
            sheet_ranges = wb['Registry']
	    row = 2
	    while True:
                desc = fix(sheet_ranges['A'+str(row)].value)
                type = str(sheet_ranges['B'+str(row)].value).strip()
                color =  str(sheet_ranges['C'+str(row)].value).strip()
	        fouo = str(sheet_ranges['D'+str(row)].value).strip()
                sighted = str(sheet_ranges['E'+str(row)].value).strip()
	        killphase=(str(sheet_ranges['F'+str(row)].value).strip()).upper()
	        hive =  str(sheet_ranges['G'+str(row)].value).strip()
	        key =  str(sheet_ranges['H'+str(row)].value).strip()
	        name =  str(sheet_ranges['I'+str(row)].value).strip()
	        data =  str(sheet_ranges['J'+str(row)].value).strip()
	        if (self.isNull(desc)) and (self.isNull(type)) and \
                    (self.isNull(color)) and (self.isNull(fouo)) and \
                    (self.isNull(sighted)) and (self.isNull(killphase)) and \
                    (self.isNull(hive)) and (self.isNull(key)) and \
                    (self.isNull(name)) and (self.isNull(data)):
	                break
                row = row + 1
                indicators.append(self.doRegistry(desc, type, color, fouo,\
                    sighted, killphase, hive, key, name, data))
            print "Processed ",row-2," Registry indicators"
        except KeyError:
	    print "Registry Sheet does not exist!"

        try:
            sheet_ranges = wb['Network Connection']
	    row = 2
	    while True:
                desc = fix(sheet_ranges['A'+str(row)].value)
                type = str(sheet_ranges['B'+str(row)].value).strip()
                color =  str(sheet_ranges['C'+str(row)].value).strip()
	        fouo = str(sheet_ranges['D'+str(row)].value).strip()
                sighted = str(sheet_ranges['E'+str(row)].value).strip()
	        killphase=(str(sheet_ranges['F'+str(row)].value).strip()).upper()
	        source =  str(sheet_ranges['G'+str(row)].value).strip()
	        sspoofed =  str(sheet_ranges['H'+str(row)].value).strip()
	        sport =  str(sheet_ranges['I'+str(row)].value).strip()
	        sproto =  str(sheet_ranges['J'+str(row)].value).strip()
	        dest =  str(sheet_ranges['K'+str(row)].value).strip()
	        dspoofed =  str(sheet_ranges['L'+str(row)].value).strip()
	        dport =  str(sheet_ranges['M'+str(row)].value).strip()
	        dproto =  str(sheet_ranges['N'+str(row)].value).strip()
	        if (self.isNull(desc)) and (self.isNull(type)) and \
                    (self.isNull(color)) and (self.isNull(fouo)) and \
                    (self.isNull(sighted)) and (self.isNull(killphase)) and \
                    (self.isNull(source)) and (self.isNull(sspoofed)) and \
                    (self.isNull(sport)) and (self.isNull(sproto)) and \
                    (self.isNull(dest)) and (self.isNull(dspoofed)) and \
                    (self.isNull(dport)) and (self.isNull(dproto)):
	                break
                row = row + 1
                indicators.append(self.doNetConn(desc, type, color, fouo,\
                    sighted, killphase, source, sspoofed, sport, sproto,\
                    dest, dspoofed,dport,dproto))
            print "Processed ",row-2," Network Connection indicators"
        except KeyError:
	    print "Network Connection Sheet does not exist!"

	return indicators

if __name__ == "__main__":

    sh = excel2stix()
    print "excel2stix "+sh.getVersion()
    print "Working, please wait!"
    if (sh.DEBUG):
        print "...STIX Version "+sh.getStixVersion()
        print "...CYBOX Version "+sh.getCyboxVersion()

    if len(sys.argv) == 1:
        print __doc__
    else:
        thepath, theext = os.path.splitext(sys.argv[1])
        outputFilename = thepath+".xml"
        myStix = {"version":"1.1.1"}
	myStix["ttps"] = sh.setKillChains()
        myStix["id"] = outputFilename
        myStix["timestamp"] = time.strftime("%Y-%m-%dT%H:%M:%S.000000Z")

        if (sh.DEBUG):
            print "...Opening excel file ",sys.argv[1]," please wait"
        warnings.filterwarnings("ignore")   #MKPMKP
        filename = sys.argv[1]
        if os.path.isfile(filename) and os.access(filename, os.R_OK):
            wb = openpyxl.load_workbook(filename=sys.argv[1])
        else:
            print "Can't open the excel file '"+filename+"'"
            sys.exit(0)

        try:
            sheet_ranges = wb['Main']
	    row = 2
            dateStr = str(sheet_ranges['A'+str(row)].value).strip()
	    titleStr = str(sheet_ranges['B'+str(row)].value).strip()
	    intentStr = str(sheet_ranges['C'+str(row)].value).strip()
	    colorStr = str(sheet_ranges['D'+str(row)].value).strip()
	    fouoStr =  str(sheet_ranges['E'+str(row)].value).strip()
	    descStr = (sheet_ranges['F'+str(row)].value)
	    namespace_URL = str(sheet_ranges['G'+str(row)].value).strip()
	    namespace_Tag = str(sheet_ranges['H'+str(row)].value).strip()

            if descStr is not None:
                descStr = fix(descStr)
	    header = getHeader(dateStr, titleStr, intentStr, colorStr,\
                fouoStr, descStr)
            myStix["stix_header"] = header
        except KeyError:
	    print "Main Sheet does not exist!"

        if (sh.DEBUG):
            print "...Got STIX Header"

	indicators = sh.getIndicators(wb)
	myStix["indicators"] = indicators
        if (sh.DEBUG):
            print "...Got STIX Indicators"

        pp = pprint.PrettyPrinter(indent=1)
        # MKPMKP
        if (sh.DEBUG):
	    print pp.pprint(myStix)

        # Change all id attributes
        tag_ns = {namespace_URL:namespace_Tag}
        cybox.utils.set_id_namespace(cybox.utils.Namespace(
            'http://www.us-cert.gov/'+namespace_Tag.lower(),namespace_Tag))

        stix.utils.set_id_namespace(tag_ns)

        p = stix.core.STIXPackage.from_dict(myStix)
        buffer = p.to_xml()

        #===========================================
        # Bug fix 20150619-03 Marlon's fix till Mitre updates this
        buffer = buffer.replace('<LinkObj:Properties xsi:type="LinkObj:LinkObjectType"',
            '<cybox:Properties xsi:type="LinkObj:LinkObjectType"')
        buffer = buffer.replace('</LinkObj:Properties>','</cybox:Properties>')

        # Use CISCP namespace for Benign or Compromised
        if buffer.find('indicator:Type>Benign') or buffer.find('indicator:Type>Compromised'):
            buffer = buffer.replace('cert.gov/'+namespace_Tag.lower()+'"',
                'cert.gov/'+namespace_Tag.lower()+'\n'    'xmlns:CISCP="http://us-cert.gov/ciscp"')
            buffer = buffer.replace('stix_core.xsd"',
                'stix_core.xsd\n    http://us-cert.gov/ciscp http://www.us-cert.gov/sites/default/files/STIX_Namespace/ciscp_vocab_v1.1.1.xsd"\n    ')
            buffer = buffer.replace('indicator:Type>Benign<',
                'indicator:Type xsi:type="CISCP:IndicatorTypeVocab-0.0">Benign<')
            buffer = buffer.replace('indicator:Type>Compromised<',
                'indicator:Type xsi:type="CISCP:IndicatorTypeVocab-0.0">Compromised<')
        buffer = "<!-- Generated by excel2stix "+sh.getVersion()+" on "+\
            time.strftime("%m/%d/%Y")+" -->\n" + buffer
        #===========================================
        if (sh.DEBUG):
            print "...Writing file ",outputFilename
        f = open(outputFilename,"w")
        f.write(buffer)

        f.close()

	print "Complete!"
