import ppc_piling_schedule

output_files = "D:/Python/engineering_scripts/output/PPC Piling Schedule.xlsx"
ppc_schedule = ppc_piling_schedule.return_ppc_piling_schedule()
ppc_schedule.to_excel(output_files, sheet_name="PPC Piling Schedule", index=False)