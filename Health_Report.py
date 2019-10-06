from Queries import DBQuery
# @ todo:make sure that 021 works with 1 or 2 rows - check num rows

report = DBQuery('Presentation/IBF_AUTOMATED_MS_Report.pptx')

report.ibs_kpi_020()  # done
report.ibs_kpi_021()  # done
report.ibs_ms_001() # table only
report.ibs_ms_003()
report.coverpage()

final_report = report
print(report)

# @todo: Mapping of IBF errors
# @todo: Email content needs to be sent
