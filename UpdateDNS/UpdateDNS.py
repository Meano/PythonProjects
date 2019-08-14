#! /usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import json
import re
import traceback
from aliyunsdkcore.client import AcsClient
from aliyunsdkalidns.request.v20150109 import DescribeDomainRecordsRequest
from aliyunsdkalidns.request.v20150109 import UpdateDomainRecordRequest

def FindRecord(domainrecords, rrname):
    for record in domainrecords:
        if record["RR"] == rrname:
            return record
    return None

if __name__ == '__main__':
    if len(sys.argv) != 6:
        print("Usage: UpdateDNS.py <ApiID> <ApiToken> <Domain> <Subdomain> <UpdateIP>")
        exit(1)

    ApiID = sys.argv[1]
    ApiToken = sys.argv[2]
    Domain = sys.argv[3]
    Subdomain = sys.argv[4]
    UpdateIP = sys.argv[5]

    if not re.match(r"^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$", UpdateIP):
        print("输入的更新IP (%s) 不合法的IP地址！" % UpdateIP)
        exit(1)

    print("将更新IP (%s) 到阿里云解析！" % UpdateIP)

    client = AcsClient(
        ApiID,
        ApiToken,
        "cn-beijing"
    )

    try:
        request = DescribeDomainRecordsRequest.DescribeDomainRecordsRequest()
        request.set_DomainName(Domain)
        request.set_PageNumber(1)
        request.set_PageSize(100)

        response = client.do_action_with_exception(request).decode("utf-8")
        DomainRecords = json.loads(response)["DomainRecords"]["Record"]

        RecordRRList = [ Subdomain ]

        for RRName in RecordRRList:
            Record = FindRecord(DomainRecords, RRName)
            if Record == None:
                print("未找到记录!")
                exit(1)
            if Record["Value"] == UpdateIP:
                print("记录值与更新值相同,不进行更新!")
                continue
            request = UpdateDomainRecordRequest.UpdateDomainRecordRequest()
            request.set_RecordId(Record["RecordId"])
            request.set_RR(Record["RR"])
            request.set_Type(Record["Type"])
            request.set_Value(UpdateIP)
            response = client.do_action_with_exception(request)
            print("记录%s成功更新为%s" % (Record["RR"], UpdateIP))
    except:
        traceback.print_exc()
        print("在进行API请求时发生了错误！即将退出！")
        exit(1)
    exit(0)




