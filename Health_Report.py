import sys
import os
from Queries import DBQuery
import argparse
from configparser import ConfigParser
import json
# @ todo:add a global prs variable
# @ todo:make sure that 021 works with 1 or 2 rows - check num rows

report = DBQuery('Presentation/IBF_AUTOMATED_MS_Report.pptx')

report.ibs_kpi_020()  # done
report.ibs_kpi_021()  # done
report.ibs_ms_001() # table only
report.ibs_ms_003()
report.coverpage()

final_report = report
print(report)

# # contract = Interface(w3)
# # functionsJson = contract.retrieveFunctions()
# print(args.input)
# contract.get_instance()
# newGreeting ='Howdy', 'hi'
# event = "GreetingChange"
#
# changeGreeting = contract.transact('setGreetingTwo', *args.input)
# # newGreeting = "I is a clever contract"
# # changeGreetingEvent = contract.transact('setGreeting', newGreeting, event='GreetingChange')
# getGreeting = contract.call('greet')
# getGreeting2 = contract.call('greetTwo')
# print(getGreeting)
# print(getGreeting2)
# print("done")

# @todo: Check if query is run on a monday or any other day
# @todo: Mapping of IBF errors
# @todo: Email content needs to be sent
# @todo: Clean up code
# @todo: Push to git